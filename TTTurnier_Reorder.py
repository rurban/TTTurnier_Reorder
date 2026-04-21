#!/usr/bin/env python3
"""
TTTurnier_Reorder - Group draw generator for TTTurnier tournaments.

Reads a registration export (.xls SpreadsheetML / .fods ODF) or the .mdb
database directly, assigns players to optimally seeded groups, resolves
same-club conflicts, writes an HTML preview, and optionally writes the
initial group layout back to the .mdb (tbl_Gruppen + tbl_Tabelle).

Usage:
    # From Excel (HTML output only):
    python TTTurnier_Reorder.py [Anmeldung.xls|Anmeldung.fods] [-o output.html]

    # From / to MDB (HTML + group setup in database):
    python TTTurnier_Reorder.py database.mdb [--class CLASSID] [-v] [-n]

    # Excel + write to a separate MDB:
    python TTTurnier_Reorder.py Anmeldung.xls --mdb database.mdb [--class 72]

Requires:
    Unix   : mdbtools >= 1.0.0 (mdb-export, mdb-sql), lxml
    Windows: pyodbc (pip install pyodbc), lxml
"""

import argparse
import csv
import html as _html_mod
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

try:
    from lxml import etree as ET

    _LXML = True
except ImportError:
    import xml.etree.ElementTree as ET

    _LXML = False

if platform.system() == "Windows":
    try:
        import pyodbc
    except ImportError:
        pyodbc = None

# ── shared utilities from KO reorder module (inline fallbacks if absent) ──────

try:
    from TTTurnier_KO_Reorder import recommend_group_count, mdb_export, mdb_sql
except ImportError:
    GROUP_COUNT_THRESHOLDS = [
        (16, 4, "4 groups of 2-4, bracket 8,  0 byes"),
        (32, 8, "8 groups of 3-4, bracket 16, 0 byes"),
        (48, 12, "12 groups of 3-4, bracket 32, 8 byes for top seeds"),
        (56, 12, "12 groups of 4-5, bracket 32, 8 byes for top seeds"),
        (80, 16, "16 groups of 4-5, bracket 32, 0 byes"),
        (160, 32, "32 groups of 4-5, bracket 64, 0 byes"),
    ]

    def recommend_group_count(n):
        for max_p, g, note in GROUP_COUNT_THRESHOLDS:
            if n <= max_p:
                return g, note
        return GROUP_COUNT_THRESHOLDS[-1][1], GROUP_COUNT_THRESHOLDS[-1][2]

    def mdb_export(mdb_file, table):
        try:
            r = subprocess.run(
                ["mdb-export", mdb_file, table],
                capture_output=True,
                text=True,
                check=True,
            )
            return list(csv.DictReader(r.stdout.splitlines()))
        except subprocess.CalledProcessError as e:
            print(f"mdb-export error ({table}): {e}", file=sys.stderr)
            return []

    def mdb_sql(mdb_file, sql_text):
        try:
            proc = subprocess.Popen(
                ["mdb-sql", "-p", mdb_file],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            out, _ = proc.communicate(input=sql_text)
            return out
        except Exception as e:
            print(f"mdb-sql error: {e}", file=sys.stderr)
            return ""


# ── XML namespace helpers ─────────────────────────────────────────────────────

_NS_SS = "urn:schemas-microsoft-com:office:spreadsheet"  # SpreadsheetML
_NS_OT = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"  # ODF table
_NS_OTX = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"  # ODF text


def _ss(tag):
    return f"{{{_NS_SS}}}{tag}"


def _ot(tag):
    return f"{{{_NS_OT}}}{tag}"


def _otx(tag):
    return f"{{{_NS_OTX}}}{tag}"


# ── XML row readers ───────────────────────────────────────────────────────────


def _row_xls(row):
    """Values from a SpreadsheetML <Row>; honours ss:Index for sparse cells."""
    out = []
    cur = 0
    for cell in row.iter(_ss("Cell")):
        idx = cell.get(_ss("Index"))
        if idx:
            cur = int(idx) - 1
        data = cell.find(_ss("Data"))
        val = (data.text or "") if data is not None else ""
        if cur >= len(out):
            out.extend([""] * (cur - len(out) + 1))
        out[cur] = val
        cur += 1
    return out


def _row_fods(row):
    """Values from an ODF table-row; expands number-columns-repeated."""
    out = []
    for cell in row.iter(_ot("table-cell")):
        repeat = int(cell.get(_ot("number-columns-repeated"), 1))
        tp = cell.find(_otx("p"))
        val = tp.text if tp is not None else None
        if val is None and repeat > 30:  # cap ODF trailing-column padding
            repeat = 1
        out.extend([val] * repeat)
    return out


# ── Spreadsheet parsers ───────────────────────────────────────────────────────


def _read_file(path):
    """Return list of (sheet_name, rows, fmt) for every worksheet in path."""
    path = Path(path)
    fmt = path.suffix.lower()
    tree = ET.parse(str(path))
    root = tree.getroot()

    if fmt in (".xls", ".xml"):
        return [
            (ws.get(_ss("Name"), ""), list(ws.iter(_ss("Row"))), "xls")
            for ws in root.iter(_ss("Worksheet"))
        ]
    elif fmt == ".fods":
        return [
            (t.get(_ot("name"), ""), list(t.iter(_ot("table-row"))), "fods")
            for t in root.iter(_ot("table"))
        ]
    else:
        raise ValueError(f"Unsupported XML format: {fmt}")


def parse_categories(sheets):
    """
    Convert raw worksheet rows into a list of category dicts.
    Each category has keys: name, title, players (list of player dicts).
    """
    categories = []
    for name, rows, fmt in sheets:
        if name == "Turnieranmeldungen" or len(rows) < 4:
            continue
        fn = _row_xls if fmt == "xls" else _row_fods

        title = (fn(rows[0]) + [""])[0] or name
        hdrs = fn(rows[1])
        col = {h: i for i, h in enumerate(hdrs) if h}

        ci_name = col.get("Nachname", 0)
        ci_vname = col.get("Vorname", 1)
        ci_club = col.get("Verein", 4)
        ci_pz = col.get("LivePZ", 10)

        players = []
        for r in rows[2:]:
            v = fn(r)

            def _c(i):
                return (v[i] or "").strip() if i < len(v) else ""

            if not _c(ci_name):
                continue
            raw_pz = _c(ci_pz)
            pz = int("".join(c for c in raw_pz if c.isdigit()) or 0)
            players.append(
                {
                    "nachname": _c(ci_name),
                    "vorname": _c(ci_vname),
                    "verein": _c(ci_club),
                    "livepz": pz,
                    "pid": None,
                }
            )

        if len(players) > 8:
            categories.append({"name": name, "title": title, "players": players})

    return categories


# ── Group assignment ──────────────────────────────────────────────────────────


def assign_groups(players, m):
    """
    Sort players by LivePZ (desc), snake-assign to m groups.

    Returns (groups, max_size).
    Each group is a list of player dicts with added keys:
      orig_group, orig_rank, moved, is_bye.
    All groups are padded to max_size with bye placeholders so that
    tta_iAV / tta_iLosPos counts are uniform across groups.
    """
    players = sorted(players, key=lambda p: p["livepz"], reverse=True)
    n = len(players)
    base = n // m
    extra = n % m
    group_sizes = [base + (1 if i < extra else 0) for i in range(m)]
    max_size = base + (1 if extra else 0)

    groups = [[] for _ in range(m)]
    pi = 0
    for slot in range(max_size):
        for gi in range(m):
            if slot >= group_sizes[gi]:
                continue
            groups[gi].append(
                {
                    **players[pi],
                    "orig_group": gi,
                    "orig_rank": pi,
                    "moved": False,
                }
            )
            pi += 1

    # Pad shorter groups with bye slots
    for gi in range(m):
        while len(groups[gi]) < max_size:
            groups[gi].append(
                {
                    "nachname": "",
                    "vorname": "",
                    "verein": "",
                    "livepz": 0,
                    "pid": "-1",
                    "orig_group": gi,
                    "orig_rank": -1,
                    "moved": False,
                    "is_bye": True,
                }
            )

    return groups, max_size


def resolve_club_conflicts(groups, m):
    """
    Iteratively swap players to eliminate same-club collisions.
    Each swap picks the partner from another group that minimises
    |LivePZ difference|.  Leaves unsolvable conflicts (club > m players) in place.
    """
    changed = True
    max_iters = m * sum(len(g) for g in groups) + 10
    iters = 0

    while changed and iters < max_iters:
        changed = False
        iters += 1

        for gi in range(m):
            real = [p for p in groups[gi] if not p.get("is_bye")]
            vc = {}
            for p in real:
                vc[p["verein"]] = vc.get(p["verein"], 0) + 1

            for verein, cnt in vc.items():
                if cnt < 2:
                    continue
                # Move the weakest (last in group-order) duplicate
                dup_idxs = [
                    k
                    for k, p in enumerate(groups[gi])
                    if p.get("verein") == verein and not p.get("is_bye")
                ]
                move_ki = dup_idxs[-1]
                mp = groups[gi][move_ki]

                best = {"gj": -1, "pj": -1, "cost": float("inf")}

                for gj in range(m):
                    if gj == gi:
                        continue
                    for pj, cp in enumerate(groups[gj]):
                        if cp.get("is_bye") or cp["verein"] == verein:
                            continue

                        # Simulate swap: neither group must gain a new collision
                        def _vc_after(grp, swap_k, new_p):
                            vc2 = {}
                            for k2, p2 in enumerate(grp):
                                p2_ = new_p if k2 == swap_k else p2
                                if not p2_.get("is_bye"):
                                    vc2[p2_["verein"]] = vc2.get(p2_["verein"], 0) + 1
                            return vc2

                        if _vc_after(groups[gi], move_ki, cp).get(cp["verein"], 0) > 1:
                            continue
                        if _vc_after(groups[gj], pj, mp).get(mp["verein"], 0) > 1:
                            continue

                        cost = abs(mp["livepz"] - cp["livepz"])
                        if cost < best["cost"]:
                            best = {"gj": gj, "pj": pj, "cost": cost}

                if best["gj"] >= 0:
                    gj = best["gj"]
                    groups[gi][move_ki], groups[gj][best["pj"]] = (
                        groups[gj][best["pj"]],
                        groups[gi][move_ki],
                    )
                    groups[gi][move_ki]["moved"] = True
                    groups[gj][best["pj"]]["moved"] = True
                    changed = True

    return groups


# ── HTML output ───────────────────────────────────────────────────────────────

_CSS = """\
body  { font-family: Arial, sans-serif; font-size: 12px; margin: 16px; }
h1    { font-size: 15px; margin: 0 0 4px; padding-bottom: 4px; border-bottom: 2px solid #444; }
.meta { font-size: 11px; color: #666; margin: 0 0 12px; }
.wrap { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 32px; page-break-after: always; }
.grp  { border: 1px solid #aaa; border-radius: 3px; min-width: 200px; }
.ghdr { background: #d4d4d4; font-weight: bold; text-align: center; padding: 3px 6px; }
table { border-collapse: collapse; width: 100%; }
td    { padding: 2px 6px; border-top: 1px solid #e8e8e8; white-space: nowrap; }
.nr   { color: #aaa; text-align: right; width: 18px; font-size: 10px; }
.pz   { text-align: right; color: #555; width: 36px; }
.red  { color: red; }
@media print { .wrap { page-break-after: always; } }"""


def write_html(categories, out_file):
    lines = [
        "<!DOCTYPE html>",
        '<html lang="de">',
        "<head>",
        '<meta charset="UTF-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        "<title>Turniergruppen</title>",
        f"<style>{_CSS}</style>",
        "</head><body>",
    ]

    for cat in categories:
        groups = cat["groups"]
        m = cat["m"]
        n = sum(1 for g in groups for p in g if not p.get("is_bye"))
        moved_n = sum(1 for g in groups for p in g if p.get("moved"))

        lines.append(f"<h1>{_html_mod.escape(cat['title'])}</h1>")
        meta = f"{n} Spieler &bull; {m} Gruppen"
        if moved_n:
            meta += f' &bull; <span class="red">{moved_n} umgesetzt</span>'
        lines.append(f'<p class="meta">{meta}</p>')
        lines.append('<div class="wrap">')

        for gi, group in enumerate(groups):
            real = sorted(
                [p for p in group if not p.get("is_bye")],
                key=lambda p: p["livepz"],
                reverse=True,
            )
            lines.append(
                f'<div class="grp"><div class="ghdr">Gruppe {gi + 1}</div><table>'
            )
            for rank, p in enumerate(real, 1):
                cls = "red" if p.get("moved") else ""
                name_str = _html_mod.escape(f"{p['nachname']}, {p['vorname']}")
                club_str = _html_mod.escape(p["verein"])
                pz_str = str(p["livepz"]) if p["livepz"] else ""
                lines.append(
                    f"<tr>"
                    f'<td class="nr">{rank}</td>'
                    f'<td class="{cls}">{name_str}</td>'
                    f'<td class="{cls}">{club_str}</td>'
                    f'<td class="pz {cls}">{pz_str}</td>'
                    f"</tr>"
                )
            lines.append("</table></div>")

        lines.append("</div>")

    lines.append("</body></html>")
    with open(out_file, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    print(f"Written: {out_file}")


# ── MDB reader ────────────────────────────────────────────────────────────────


def read_players_from_mdb(mdb_file, klassen_id, verbose=False):
    """
    Return player dicts for the given class from tbl_Spieler,
    filtered via cor_Spieler_Anmeldung.
    """
    spieler_rows = mdb_export(mdb_file, "tbl_Spieler")
    csa_rows = mdb_export(mdb_file, "cor_Spieler_Anmeldung")

    registered_pids = {
        r["csa_refSpielerID"]
        for r in csa_rows
        if r.get("csa_refKlassenID") == str(klassen_id)
    }

    players = []
    for p in spieler_rows:
        pid = p.get("ts_ID")
        if pid not in registered_pids:
            continue
        try:
            livepz = int(p.get("ts_sSpielstaerke") or 0)
        except ValueError:
            livepz = 0
        players.append(
            {
                "nachname": p.get("ts_sNachname", ""),
                "vorname": p.get("ts_sVorname", ""),
                "verein": p.get("ts_sVereinName", ""),
                "livepz": livepz,
                "pid": pid,
            }
        )

    if verbose:
        print(f"  Class {klassen_id}: {len(players)} players loaded from MDB")
    return players


# ── MDB writer ────────────────────────────────────────────────────────────────


def _v(val):
    """Format a value for SQL: None → NULL, strings → single-quoted."""
    if val is None:
        return "NULL"
    if isinstance(val, str):
        return f"'{val.replace(chr(39), chr(39) * 2)}'"
    return str(val)


def _execute_sql(mdb_file, stmts, dry_run, verbose):
    """Run a list of SQL statements via mdb-sql (Unix) or pyodbc (Windows)."""
    if not stmts:
        return
    if dry_run:
        for s in stmts:
            print(f"  SQL: {s}")
        return

    if platform.system() == "Windows" and pyodbc:
        conn_str = (
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={os.path.abspath(str(mdb_file))};"
        )
        try:
            conn = pyodbc.connect(conn_str)
            cur = conn.cursor()
            for s in stmts:
                cur.execute(s)
            conn.commit()
            cur.close()
            conn.close()
        except Exception as e:
            print(f"pyodbc error: {e}", file=sys.stderr)
        return

    # Unix: batch all statements, separated by newlines
    block = "\n".join(s.rstrip(";") + ";" for s in stmts) + "\n"
    if verbose:
        print(block[:800])
    mdb_sql(str(mdb_file), block)


def write_groups_to_mdb(categories, mdb_file, turnier_id, dry_run=False, verbose=False):
    """
    For each category:
      1. Abort if Phase 1 games already have results (safety check).
      2. Delete existing Phase 1 tbl_Spiele, tbl_Tabelle, tbl_Gruppen rows.
      3. Insert new tbl_Gruppen rows (one per group, tgr_ID auto-incremented).
      4. Insert tbl_Tabelle rows: real players + bye placeholders (-1), all
         statistics zeroed.  tta_iLosPos = LivePZ rank within group (1=best).
         tta_iAV = global sequential slot across all groups.
         tta_iLosPosChanged = 1 for players moved by club-conflict resolution.
    """
    grp_rows = mdb_export(str(mdb_file), "tbl_Gruppen")
    max_grp_id = max(
        (int(r["tgr_ID"]) for r in grp_rows if r.get("tgr_ID")),
        default=0,
    )

    for cat in categories:
        klassen_id = cat.get("klassen_id")
        if not klassen_id:
            print(
                f"  Skipping '{cat['name']}': no klassen_id resolved", file=sys.stderr
            )
            continue

        groups = cat["groups"]
        m = len(groups)

        # Safety: do not overwrite Phase 1 that already has game results
        spiele = mdb_export(str(mdb_file), "tbl_Spiele")
        played = [
            r
            for r in spiele
            if r.get("tsp_refTurnier") == str(turnier_id)
            and r.get("tsp_refKlasse") == str(klassen_id)
            and r.get("tsp_iPhase") == "1"
            and (r.get("tsp_iSatzA") or r.get("tsp_iSatzB"))
        ]
        if played:
            print(
                f"WARNING: {len(played)} Phase 1 game(s) with results found for "
                f"class {klassen_id} – skipping MDB write.",
                file=sys.stderr,
            )
            continue

        # ── delete existing Phase 1 data ───────────────────────────────────
        _execute_sql(
            str(mdb_file),
            [
                f"DELETE FROM tbl_Spiele "
                f"WHERE tsp_refTurnier={turnier_id} AND tsp_refKlasse={klassen_id} "
                f"AND tsp_iPhase=1",
                f"DELETE FROM tbl_Tabelle "
                f"WHERE tta_refTurnierID={turnier_id} AND tta_refKlassenID={klassen_id}",
                f"DELETE FROM tbl_Gruppen "
                f"WHERE tgr_refTurnierID={turnier_id} AND tgr_refKlassenID={klassen_id} "
                f"AND tgr_iPhase=1",
            ],
            dry_run,
            verbose,
        )

        # ── insert new groups and player slots ─────────────────────────────
        inserts = []
        global_slot = 1  # tta_iAV: sequential across all groups in this class

        for gi, group in enumerate(groups):
            max_grp_id += 1
            grp_id = max_grp_id
            grp_num = gi + 1

            inserts.append(
                f"INSERT INTO tbl_Gruppen "
                f"(tgr_refTurnierID, tgr_refKlassenID, tgr_iPhase, "
                f" tgr_ID, tgr_iPos, tgr_sGruppenName, tgr_refTisch) "
                f"VALUES ({turnier_id}, {klassen_id}, 1, "
                f"{grp_id}, {grp_num}, '{grp_num}', NULL)"
            )

            # Order real players by LivePZ desc (→ tta_iLosPos 1 = best seed).
            real = sorted(
                [p for p in group if not p.get("is_bye")],
                key=lambda p: p["livepz"],
                reverse=True,
            )
            byes = [p for p in group if p.get("is_bye")]
            slots = real + byes  # byes fill trailing slot(s)

            for slot_pos, p in enumerate(slots):
                is_bye = p.get("is_bye", False)
                los_pos = slot_pos + 1  # position within group
                av = global_slot
                global_slot += 1

                pid = "-1" if is_bye else p["pid"]
                platz = "NULL" if is_bye else (slot_pos + 1)
                ref_sp2 = -1 if is_bye else 0  # tta_refSpieler2
                los_chg = "NULL" if is_bye else (1 if p.get("moved") else 0)
                buch = "NULL" if is_bye else 0  # tta_iBuchholz1/2

                inserts.append(
                    f"INSERT INTO tbl_Tabelle "
                    f"(tta_refTurnierID, tta_refKlassenID, tta_refGruppenID, "
                    f" tta_refSpieler, tta_iPlatz, "
                    f" tta_iST, tta_iSpielePlus, tta_iSpieleMinus, "
                    f" tta_iSaetzePlus, tta_iSaetzeMinus, tta_iBalleDif, "
                    f" tta_refSpieler2, tta_iLosPos, "
                    f" tta_iVGSpielePlus, tta_iVGSpieleMinus, "
                    f" tta_iVGSaetzePlus, tta_iVGSaetzeMinus, tta_iVGBalleDif, "
                    f" tta_iAV, tta_iLosPosChanged, tta_iBuchholz1, tta_iBuchholz2) "
                    f"VALUES "
                    f"({turnier_id}, {klassen_id}, {grp_id}, "
                    f" {pid}, {platz}, "
                    f" 0, 0, 0, 0, 0, 0, "
                    f" {ref_sp2}, {los_pos}, "
                    f" NULL, NULL, NULL, NULL, NULL, "
                    f" {av}, {los_chg}, {buch}, {buch})"
                )

        _execute_sql(str(mdb_file), inserts, dry_run, verbose)

        real_count = sum(1 for g in groups for p in g if not p.get("is_bye"))
        print(
            f"{'[dry-run] ' if dry_run else ''}"
            f"Class {klassen_id}: {m} groups, "
            f"{real_count} players, {global_slot - 1 - real_count} byes, "
            f"{global_slot - 1} total slots written"
        )


# ── main ──────────────────────────────────────────────────────────────────────


def main():
    parser = argparse.ArgumentParser(
        description="TTTurnier group draw: sort players into seeded groups, "
        "resolve club conflicts, write HTML and/or .mdb."
    )
    parser.add_argument(
        "input", nargs="?", help=".xls, .fods, or .mdb file (auto-detected if omitted)"
    )
    parser.add_argument(
        "--mdb",
        help="Write group assignments to this .mdb (implied when "
        "input is already .mdb)",
    )
    parser.add_argument(
        "--class",
        dest="klassen_id",
        type=int,
        help="Klassen-ID to process (required when input is .mdb; "
        "auto when only one class has > 8 players)",
    )
    parser.add_argument("-o", "--output", help="HTML output file path")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument(
        "-n",
        "--dry-run",
        action="store_true",
        help="Print SQL statements but do not modify the .mdb",
    )
    args = parser.parse_args()

    exe_dir = Path(__file__).parent

    # ── locate input file ─────────────────────────────────────────────────
    in_path = Path(args.input) if args.input else None
    if not in_path:
        for pat in ["Anmeldung*.fods", "Anmeldung*.xls", "*.mdb"]:
            cands = sorted(exe_dir.glob(pat))
            if cands:
                in_path = cands[0]
                break

    if not in_path or not in_path.exists():
        print(
            "No input file found.  Provide an Anmeldung.xls/fods or a .mdb file.",
            file=sys.stderr,
        )
        sys.exit(1)

    fmt = in_path.suffix.lower()

    # ── determine MDB target (if any) ────────────────────────────────────
    mdb_file = Path(args.mdb) if args.mdb else (in_path if fmt == ".mdb" else None)

    # ── HTML output path ──────────────────────────────────────────────────
    out_file = args.output or str(
        exe_dir / (("Gruppen" if fmt == ".mdb" else in_path.stem) + ".html")
    )

    # ── load categories ───────────────────────────────────────────────────
    turnier_id = None
    categories = []

    if fmt == ".mdb":
        turniere = mdb_export(str(in_path), "tbl_Turniere")
        if not turniere:
            print("No tournament found in MDB.", file=sys.stderr)
            sys.exit(1)
        turnier_id = int(turniere[0]["tt_ID"])

        klassen = mdb_export(str(in_path), "tbl_Klassen")
        for k in klassen:
            kid = int(k["tkl_ID"])
            if args.klassen_id and kid != args.klassen_id:
                continue
            players = read_players_from_mdb(str(in_path), kid, args.verbose)
            if len(players) <= 8:
                continue
            categories.append(
                {
                    "name": k.get("tkl_sKlassenname", str(kid)),
                    "title": k.get("tkl_sKlassenname", str(kid)),
                    "klassen_id": kid,
                    "players": players,
                }
            )

    elif fmt in (".xls", ".xml", ".fods"):
        try:
            sheets = _read_file(in_path)
        except Exception as e:
            print(f"Failed to parse {in_path}: {e}", file=sys.stderr)
            sys.exit(1)
        categories = parse_categories(sheets)

        # If --mdb is given, try to match sheet names to class names in the DB
        if mdb_file and mdb_file.exists():
            klassen = mdb_export(str(mdb_file), "tbl_Klassen")
            for cat in categories:
                for k in klassen:
                    if cat["name"] == k.get("tkl_sKlassenname", ""):
                        cat["klassen_id"] = int(k["tkl_ID"])
                        break

    else:
        print(f"Unsupported input format: {fmt}", file=sys.stderr)
        sys.exit(1)

    if not categories:
        print("No categories with > 8 players found.", file=sys.stderr)
        sys.exit(1)

    # ── assign groups and resolve conflicts ───────────────────────────────
    for cat in categories:
        n = len(cat["players"])
        m, note = recommend_group_count(n)
        print(f"\n{cat['name']}: {n} players → {m} groups ({note})")

        groups, max_size = assign_groups(cat["players"], m)
        groups = resolve_club_conflicts(groups, m)

        moved_n = sum(1 for g in groups for p in g if p.get("moved"))
        byes_n = sum(1 for g in groups for p in g if p.get("is_bye"))
        print(f"  max_size={max_size}, byes={byes_n}, moved_by_club_rule={moved_n}")

        if args.verbose:
            for gi, g in enumerate(groups):
                real = [p for p in g if not p.get("is_bye")]
                real_s = sorted(real, key=lambda p: p["livepz"], reverse=True)
                dups = [
                    v
                    for v in {p["verein"] for p in real}
                    if sum(1 for p in real if p["verein"] == v) > 1
                ]
                tag = f" ** CLUB CONFLICT: {dups}" if dups else ""
                print(
                    f"  G{gi + 1}: "
                    + ", ".join(
                        f"{'*' if p.get('moved') else ''}{p['nachname']}({p['livepz']})"
                        for p in real_s
                    )
                    + tag
                )

        cat.update({"groups": groups, "m": m, "max_size": max_size})

    # ── HTML output ───────────────────────────────────────────────────────
    write_html(categories, out_file)

    # ── MDB write ─────────────────────────────────────────────────────────
    if mdb_file:
        if not mdb_file.exists():
            print(f"MDB not found: {mdb_file}", file=sys.stderr)
            sys.exit(1)

        # Resolve turnier_id from the target MDB if input was Excel
        if turnier_id is None:
            turniere = mdb_export(str(mdb_file), "tbl_Turniere")
            if not turniere:
                print("No tournament found in target MDB.", file=sys.stderr)
                sys.exit(1)
            turnier_id = int(turniere[0]["tt_ID"])

        backup = mdb_file.with_suffix(mdb_file.suffix + ".bak")
        if backup.exists():
            print(f"Backup already exists: {backup} (not overwriting)")
        else:
            shutil.copy2(str(mdb_file), str(backup))
            print(f"Backed up: {backup}")

        cats_writable = [c for c in categories if c.get("klassen_id")]
        if not cats_writable:
            print(
                "No class IDs resolved for MDB write.  "
                "Use --class or ensure sheet names match tbl_Klassen.",
                file=sys.stderr,
            )
        else:
            write_groups_to_mdb(
                cats_writable,
                mdb_file,
                turnier_id,
                dry_run=args.dry_run,
                verbose=args.verbose,
            )

    # ── summary ───────────────────────────────────────────────────────────
    print(f"\nProcessed {len(categories)} category(s):")
    for cat in categories:
        moved = sum(1 for g in cat["groups"] for p in g if p.get("moved"))
        n = sum(1 for g in cat["groups"] for p in g if not p.get("is_bye"))
        print(
            f"  {cat['name']:<32}  {n:>3} players  {cat['m']:>2} groups  {moved} moved"
        )


if __name__ == "__main__":
    main()

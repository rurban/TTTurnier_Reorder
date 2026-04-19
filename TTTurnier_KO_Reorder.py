#!/usr/bin/env python3
"""
TTTurnier_KO_Reorder - Reorder Phase 2 KO pairings for TTTurnier tournaments.
Reads from .mdb database (via mdbtools on Unix or pyodbc on Windows), rearranges Round of 32, writes back.

Usage:
    python TTTurnier_KO_Reorder.py [-v] [-n] database.mdb
    python TTTurnier_KO_Reorder.py -v sem_b_2026.mdb   # verbose
    python TTTurnier_KO_Reorder.py -n sem_b_2026.mdb   # dry-run

Requires:
    Unix: mdbtools (mdb-export, mdb-sql)
    Windows: pyodbc (pip install pyodbc)
"""

import csv
import subprocess
import sys
import os
import shutil
from pathlib import Path
import argparse
import platform

# Windows-specific imports
if platform.system() == "Windows":
    try:
        import pyodbc
    except ImportError:
        pyodbc = None


def mdb_export_win(mdb_file, table):
    """Export table from .mdb on Windows using pyodbc."""
    try:
        conn_str = (
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={mdb_file};"
        )
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM [{table}]")

        # Get column names
        columns = [column[0] for column in cursor.description]

        # Fetch all rows
        rows = []
        for row in cursor.fetchall():
            rows.append(
                dict(zip(columns, [str(val) if val is not None else "" for val in row]))
            )

        cursor.close()
        conn.close()
        return rows
    except Exception as e:
        print(f"Error exporting {table} on Windows: {e}", file=sys.stderr)
        return []


def mdb_export(mdb_file, table):
    """Export table from .mdb using mdbtools on Unix or pyodbc on Windows."""
    if platform.system() == "Windows" and pyodbc:
        return mdb_export_win(mdb_file, table)

    try:
        result = subprocess.run(
            ["mdb-export", mdb_file, table], capture_output=True, text=True, check=True
        )
        reader = csv.DictReader(result.stdout.splitlines())
        return list(reader)
    except subprocess.CalledProcessError as e:
        print(f"Error exporting {table}: {e}", file=sys.stderr)
        return []


def mdb_sql_win(mdb_file, sql):
    """Execute SQL on .mdb on Windows using pyodbc."""
    if platform.system() != "Windows" or not pyodbc:
        return mdb_sql(mdb_file, sql)

    try:
        conn_str = (
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={mdb_file};"
        )
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(sql)
        conn.commit()
        affected = cursor.rowcount
        cursor.close()
        conn.close()
        return f"{affected} rows affected"
    except Exception as e:
        print(f"Error executing SQL on Windows: {e}", file=sys.stderr)
        return ""


def mdb_sql(mdb_file, sql):
    """Execute SQL on .mdb using mdb-sql on Unix or pyodbc on Windows."""
    if platform.system() == "Windows" and pyodbc:
        return mdb_sql_win(mdb_file, sql)

    try:
        proc = subprocess.Popen(
            ["mdb-sql", "-p", mdb_file],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        stdout, stderr = proc.communicate(input=sql)
        return stdout
    except Exception as e:
        print(f"Error executing SQL: {e}", file=sys.stderr)
        return ""


def parse_args():
    parser = argparse.ArgumentParser(description="Reorder Phase 2 KO pairings")
    parser.add_argument("mdb_file", nargs="?", default=None)
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("-n", "--dry-run", action="store_true")
    return parser.parse_args()


def main():
    args = parse_args()

    # Locate .mdb file
    if args.mdb_file:
        mdb_file = args.mdb_file
    else:
        # Auto-detect
        exe_dir = Path(__file__).parent
        mdb_files = list(exe_dir.glob("*.mdb"))
        if not mdb_files:
            print("No .mdb file found", file=sys.stderr)
            sys.exit(1)
        mdb_file = str(mdb_files[0])

    if not os.path.isfile(mdb_file):
        print(f"File not found: {mdb_file}", file=sys.stderr)
        sys.exit(1)

    # Backup
    mdb_dir = os.path.dirname(mdb_file) or "."
    mdb_name = os.path.basename(mdb_file)
    backup = os.path.join(mdb_dir, f"{mdb_name}.bak")

    if os.path.exists(backup):
        print(f"Backup already exists: {backup}")
    else:
        shutil.copy2(mdb_file, backup)
        print(f"Backed up: {backup}")

    verbose = args.verbose
    dry_run = args.dry_run

    # Export tables
    print(f"Reading tables from: {mdb_file}")

    players_raw = mdb_export(mdb_file, "tbl_Spieler")
    players = {}
    for p in players_raw:
        pid = p.get("ts_ID")
        if not pid:
            continue
        try:
            livepz = int(p.get("ts_sSpielstaerke", 0) or 0)
        except ValueError:
            livepz = 0
        players[pid] = {
            "livepz": livepz,
            "verein": p.get("ts_sVereinName", ""),
            "name": f"{p.get('ts_Vorname', '')} {p.get('ts_sNachname', '')}".strip(),
        }

    print(f"Loaded {len(players)} players")

    # Read group results
    tabelle_raw = mdb_export(mdb_file, "tbl_Tabelle")
    group_results = {}  # gruppe -> {1: winner_id, 2: second_id}

    for row in tabelle_raw:
        gruppe = row.get("tta_refGruppenID")
        platz = row.get("tta_iPlatz")
        spieler = row.get("tta_refSpieler")

        if not gruppe or not platz or not spieler or spieler == "-1":
            continue
        try:
            platz = int(platz)
        except ValueError:
            continue
        if platz not in (1, 2):
            continue

        if gruppe not in group_results:
            group_results[gruppe] = {}
        group_results[gruppe][platz] = spieler

    groups = sorted(group_results.keys(), key=lambda x: int(x))
    print(f"Found {len(groups)} groups")

    if verbose:
        for g in groups:
            w = group_results[g].get(1, "-1")
            s = group_results[g].get(2, "-1")
            w_name = players.get(w, {}).get("name", "?") if w != "-1" else "N/A"
            s_name = players.get(s, {}).get("name", "?") if s != "-1" else "N/A"
            print(f"  Group {g}: winner={w} ({w_name}), second={s} ({s_name})")

    # Extract winners and seconds by group order
    winners = []
    seconds = []
    for g in groups:
        winners.append(group_results[g].get(1, "-1"))
        seconds.append(group_results[g].get(2, "-1"))

    # Fixed position mapping for 16 group winners (CLAUDE.md)
    winner_pos = {
        1: 1,
        2: 32,
        3: 17,
        4: 16,
        5: 9,
        6: 24,
        7: 25,
        8: 8,
        9: 12,
        10: 21,
        11: 28,
        12: 5,
        13: 13,
        14: 20,
        15: 30,
        16: 4,
    }

    # For seconds: placed in same match pair as winner
    # Match pairs: (1,2), (3,4), ..., (31,32)
    # Winner at position p: second at p+1 if p is odd, p-1 if p is even
    second_pos = {}
    for grp in range(1, 17):
        wp = winner_pos[grp]
        if wp % 2 == 1:  # odd
            second_pos[grp] = wp + 1
        else:
            second_pos[grp] = wp - 1

    # Build position map
    position_map = {}  # position -> {pid, type, group}
    for i, g in enumerate(groups, 1):
        pid = winners[i - 1] if i - 1 < len(winners) else "-1"
        if pid and pid != "-1" and pid in players:
            position_map[winner_pos[i]] = {"pid": pid, "type": f"G{i}P1", "group": i}

        pid = seconds[i - 1] if i - 1 < len(seconds) else "-1"
        if pid and pid != "-1" and pid in players:
            position_map[second_pos[i]] = {"pid": pid, "type": f"G{i}P2", "group": i}

    print(f"Position map has {len(position_map)} players")

    # Check conflicts
    def check_conflicts(pos_map):
        conflicts = []
        for i in range(16):
            pos_a = 2 * i + 1
            pos_b = 2 * i + 2
            if pos_a in pos_map and pos_b in pos_map:
                p_a = players.get(pos_map[pos_a]["pid"])
                p_b = players.get(pos_map[pos_b]["pid"])
                if p_a and p_b and p_a["verein"] == p_b["verein"]:
                    conflicts.append(
                        {
                            "pos_a": pos_a,
                            "pos_b": pos_b,
                            "name_a": p_a["name"],
                            "name_b": p_b["name"],
                            "club": p_a["verein"],
                        }
                    )
        return conflicts

    conflicts = check_conflicts(position_map)
    print(f"Initial club conflicts: {len(conflicts)}")
    if verbose:
        for c in conflicts:
            print(
                f"  Match {c['pos_a']} vs {c['pos_b']}: {c['name_a']} vs {c['name_b']} ({c['club']})"
            )

    # Try to resolve conflicts
    if conflicts:
        # Try swapping within same match
        for c in conflicts:
            pos_a, pos_b = c["pos_a"], c["pos_b"]
            # Swap
            tmp = position_map[pos_a]
            position_map[pos_a] = position_map[pos_b]
            position_map[pos_b] = tmp

            new_conflicts = check_conflicts(position_map)
            if len(new_conflicts) < len(conflicts):
                conflicts = new_conflicts
                print(f"Resolved by swapping within match {pos_a}/{pos_b}")
                break
            else:
                # Swap back
                tmp = position_map[pos_a]
                position_map[pos_a] = position_map[pos_b]
                position_map[pos_b] = tmp

    print(f"Final club conflicts: {len(conflicts)}")

    # Print final pairings
    print("\n=== FINAL KO ROUND 1 PAIRINGS ===\n")
    print(
        f"{'Pos':>3}  {'Player A':<25} {'LivePZ':>6}  {'Player B':<25} {'LivePZ':>6}  {'Club'}"
    )
    print("-" * 95)

    for i in range(16):
        pos_a = 2 * i + 1
        pos_b = 2 * i + 2

        p_a_data = position_map.get(pos_a)
        p_b_data = position_map.get(pos_b)

        if not p_a_data or not p_b_data:
            continue

        p_a = players.get(p_a_data["pid"])
        p_b = players.get(p_b_data["pid"])

        if not p_a or not p_b:
            continue

        conflict = " ***" if p_a["verein"] == p_b["verein"] else ""
        print(
            f"{pos_a:>2}-{pos_b:<2}  {p_a['name'][:24]:<24} {p_a['livepz']:>6}  {p_b['name'][:24]:<24} {p_b['livepz']:>6}  {p_a['verein'][:20]}{conflict}"
        )

    # Read Phase 2 games
    spiele_raw = mdb_export(mdb_file, "tbl_Spiele")
    phase2_r1 = [
        r
        for r in spiele_raw
        if r.get("tsp_iPhase") == "2" and r.get("tsp_iRunde") == "1"
    ]

    if not phase2_r1:
        print("No Phase 2 Round 1 games found")
        return

    print(f"\nFound {len(phase2_r1)} Phase 2 Round 1 games")

    # Generate SQL updates
    updates = []
    for row in phase2_r1:
        try:
            pos_a = int(row.get("tsp_iPosAPlan", 0))
            pos_b = int(row.get("tsp_iPosBPlan", 0))
        except (ValueError, TypeError):
            continue

        if pos_a not in position_map or pos_b not in position_map:
            continue

        new_a = position_map[pos_a]["pid"]
        new_b = position_map[pos_b]["pid"]

        old_a = row.get("tsp_refSpielerA_1")
        old_b = row.get("tsp_refSpielerB_1")

        if old_a != new_a or old_b != new_b:
            game_id = row.get("tsp_ID")
            updates.append(
                {
                    "game_id": game_id,
                    "old_a": old_a,
                    "old_b": old_b,
                    "new_a": new_a,
                    "new_b": new_b,
                    "pos_a": pos_a,
                    "pos_b": pos_b,
                }
            )

    if not updates:
        print("No changes needed")
        return

    print(f"\n{len(updates)} games need updating")

    if dry_run:
        print("\nDry run - SQL statements:")
        for u in updates:
            print(
                f"UPDATE tbl_Spiele SET tsp_refSpielerA_1='{u['new_a']}', tsp_refSpielerB_1='{u['new_b']}' WHERE tsp_ID={u['game_id']};"
            )
    else:
        # Write updates using mdb-sql
        print("\nApplying updates...")
        for u in updates:
            sql = f"UPDATE tbl_Spiele SET tsp_refSpielerA_1='{u['new_a']}', tsp_refSpielerB_1='{u['new_b']}' WHERE tsp_ID={u['game_id']};\n"
            mdb_sql(mdb_file, sql)
            if verbose:
                print(
                    f"Updated game {u['game_id']}: "
                    f"A: {u['old_a']} -> {u['new_a']} "
                    f"({u['old_a_name']} -> {u['new_a_name']}), "
                    f"B: {u['old_b']} -> {u['new_b']} "
                    f"({u['old_b_name']} -> {u['new_b_name']})"
                )

        print("Done!")

    print(f"\nBackup available at: {backup}")


if __name__ == "__main__":
    main()

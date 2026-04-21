#!/usr/bin/env python3
"""
TTTurnier_KO_Reorder - Reorder Phase 2 KO pairings for TTTurnier tournaments.
Reads from .mdb database (via mdbtools on Unix or pyodbc on Windows),
rearranges first round of Phase 2, the KO round, writes back to the db.

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
    if not pyodbc:
        # verbose is not available here, always print error for critical dependency
        print("pyodbc not available on Windows", file=sys.stderr)
        return []

    # Ensure we have a proper string path
    if not isinstance(mdb_file, str):
        mdb_file = str(mdb_file)

    # Handle potential path issues on Windows
    try:
        # Convert to absolute path to avoid any relative path issues
        mdb_file = os.path.abspath(mdb_file)
    except Exception:
        pass  # If we can't make it absolute, use as-is

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
    except pyodbc.Error as e:
        print(
            f"Error exporting {table} on Windows (pyodbc error): {e}", file=sys.stderr
        )
        print(
            f"  Connection string was: Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_file}",
            file=sys.stderr,
        )
        return []
    except Exception as e:
        print(f"Error exporting {table} on Windows: {e}", file=sys.stderr)
        print(
            f"  Connection string was: Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_file}",
            file=sys.stderr,
        )
        return []

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
    except pyodbc.Error as e:
        print(
            f"Error exporting {table} on Windows (pyodbc error): {e}", file=sys.stderr
        )
        return []
    except Exception as e:
        print(f"Error exporting {table} on Windows: {e}", file=sys.stderr)
        return []

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
    except pyodbc.Error as e:
        print(
            f"Error exporting {table} on Windows (pyodbc error): {e}", file=sys.stderr
        )
        return []
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

    # Ensure we have a proper string path
    if not isinstance(mdb_file, str):
        mdb_file = str(mdb_file)

    # Handle potential path issues on Windows
    try:
        # Convert to absolute path to avoid any relative path issues
        mdb_file = os.path.abspath(mdb_file)
    except Exception:
        pass  # If we can't make it absolute, use as-is

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
    except pyodbc.Error as e:
        print(f"Error executing SQL on Windows (pyodbc error): {e}", file=sys.stderr)
        print(
            f"  Connection string was: Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_file}",
            file=sys.stderr,
        )
        return ""
    except Exception as e:
        print(f"Error executing SQL on Windows: {e}", file=sys.stderr)
        print(
            f"  Connection string was: Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_file}",
            file=sys.stderr,
        )
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


def standard_seeding(n):
    """Standard tournament seeding for a bracket of size n (must be a power of 2).

    Returns a list where element i (0-indexed) is the seed (1-indexed) placed at
    bracket position i+1.  Seed 1 and seed 2 meet only in the final, seeds 1/3/4
    meet only in semis, etc.

    Examples:
        standard_seeding(2)  -> [1, 2]
        standard_seeding(4)  -> [1, 4, 2, 3]
        standard_seeding(8)  -> [1, 8, 4, 5, 2, 7, 3, 6]
        standard_seeding(16) -> [1, 16, 8, 9, 4, 13, 5, 12, 2, 15, 7, 10, 3, 14, 6, 11]
    """
    if n <= 1:
        return [1]
    prev = standard_seeding(n // 2)
    result = []
    for s in prev:
        result.append(s)
        result.append(n + 1 - s)
    return result


# Threshold table: player count -> recommended group count.
# Aim for 4-5 players per group.  Group count should be a power of 2 so the
# full KO bracket (2*m) is also a power of 2 and round-1 byes are avoided.
# Non-power-of-2 counts are supported but the top seeds receive round-1 byes.
GROUP_COUNT_THRESHOLDS = [
    # (max_players, groups,  note)
    (16, 4, "4 groups of 2-4, bracket 8,  0 byes"),
    (32, 8, "8 groups of 3-4, bracket 16, 0 byes"),
    (48, 12, "12 groups of 3-4, bracket 32, 8 byes for top seeds"),
    (56, 12, "12 groups of 4-5, bracket 32, 8 byes for top seeds"),
    (80, 16, "16 groups of 4-5, bracket 32, 0 byes"),
    (160, 32, "32 groups of 4-5, bracket 64, 0 byes"),
]


def recommend_group_count(num_players):
    """Return (groups, note) for the given player count."""
    for max_p, groups, note in GROUP_COUNT_THRESHOLDS:
        if num_players <= max_p:
            return groups, note
    g, n = GROUP_COUNT_THRESHOLDS[-1][1], GROUP_COUNT_THRESHOLDS[-1][2]
    return g, n


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
        # Auto-detect: take the newest .mdb file in the mdb subdirectory
        exe_dir = Path(__file__).parent
        mdb_dir = exe_dir / "mdb"
        if not mdb_dir.is_dir():
            print("No mdb directory found", file=sys.stderr)
            sys.exit(1)
        mdb_files = list(mdb_dir.glob("*.mdb"))
        if not mdb_files:
            print("No .mdb file found in mdb directory", file=sys.stderr)
            sys.exit(1)
        # Sort by modification time, newest first
        mdb_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        mdb_file = str(mdb_files[0])

    if not os.path.isfile(mdb_file):
        print(f"File not found: {mdb_file}", file=sys.stderr)
        sys.exit(1)

    # Backup
    mdb_dir = os.path.dirname(mdb_file) or "."
    mdb_name = os.path.basename(mdb_file)
    backup = os.path.join(mdb_dir, f"{mdb_name}.bak")

    if os.path.exists(backup):
        print(f"Backup already exists: {backup} (not overwriting)")
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

    num_groups = len(groups)
    num_advancers = num_groups * 2  # winner + second from each group

    # winner_bracket: smallest power of 2 >= num_groups.
    # Full KO bracket = 2 * winner_bracket (one winner + one runner-up per match).
    winner_bracket = 1
    while winner_bracket < num_groups:
        winner_bracket *= 2
    bracket_size = winner_bracket * 2
    num_byes = bracket_size - num_advancers

    # Advisory: recommend group count for the player total visible in tbl_Spieler.
    rec_groups, rec_note = recommend_group_count(len(players))
    if num_groups != rec_groups:
        print(
            f"Note: {len(players)} players -> recommended {rec_groups} groups "
            f"({rec_note}); found {num_groups} groups."
        )

    print(
        f"Groups: {num_groups}, Advancers: {num_advancers}, "
        f"Bracket: {bracket_size}, Byes: {num_byes}"
        + (" (top seeds advance without a round-1 match)" if num_byes else "")
    )

    # Build seed -> winner position using standard tournament seeding.
    #
    # standard_seeding(winner_bracket) gives the seed at each position 1..winner_bracket.
    # Winner of seed s (= group s, ranked by initial LivePZ) goes to the A-side
    # (odd position) of the corresponding full-bracket match:
    #   winner_bracket pos k  ->  full bracket pos 2k-1  (A-side, always odd)
    #
    # Runner-up of seed s goes to the B-side (even position) of the mirror match
    # in the opposite half of the full bracket, so winner and runner-up from the
    # same group can never meet in round 1.
    #
    # Formula (half = bracket_size // 2):
    #   If full_pos <= half  ->  second_pos = full_pos + half + 1   (lower half, even)
    #   If full_pos >  half  ->  second_pos = full_pos - half + 1   (upper half, even)
    #
    # Seeds beyond num_groups are phantom (no player); their positions become byes.

    seeding = standard_seeding(winner_bracket)  # list: pos -> seed (1-indexed)
    seed_to_wpos = {}  # seed -> winner-bracket position
    for pos_idx, seed in enumerate(seeding):
        seed_to_wpos[seed] = pos_idx + 1

    half = bracket_size // 2
    winner_pos = {}  # seed -> full-bracket position (odd)
    second_pos = {}  # seed -> full-bracket position (even, opposite half)

    for seed, wpos in seed_to_wpos.items():
        full_pos = 2 * wpos - 1  # A-side, odd
        winner_pos[seed] = full_pos
        if full_pos <= half:
            second_pos[seed] = full_pos + half + 1
        else:
            second_pos[seed] = full_pos - half + 1

    if verbose:
        print(
            f"Winner positions (seeds 1..{num_groups}): "
            f"{[winner_pos[s] for s in range(1, num_groups + 1)]}"
        )
        print(
            f"Second positions (seeds 1..{num_groups}): "
            f"{[second_pos[s] for s in range(1, num_groups + 1)]}"
        )
        bye_seeds = [s for s in range(num_groups + 1, winner_bracket + 1)]
        if bye_seeds:
            bye_w = [winner_pos[s] for s in bye_seeds]
            bye_s = [second_pos[s] for s in bye_seeds]
            print(
                f"Bye positions (seeds {bye_seeds[0]}..{bye_seeds[-1]}): "
                f"winner slots {bye_w}, second slots {bye_s}"
            )

    # Build position map. Only real groups (seeds 1..num_groups) get players.
    # Phantom-seed positions stay absent -> byes.
    position_map = {}  # full-bracket position -> {pid, type, group}
    for i, g in enumerate(groups, 1):
        seed = i
        w_pid = winners[seed - 1] if seed - 1 < len(winners) else "-1"
        s_pid = seconds[seed - 1] if seed - 1 < len(seconds) else "-1"

        if w_pid and w_pid != "-1" and w_pid in players:
            position_map[winner_pos[seed]] = {
                "pid": w_pid,
                "type": f"G{seed}P1",
                "group": seed,
            }
        if s_pid and s_pid != "-1" and s_pid in players:
            position_map[second_pos[seed]] = {
                "pid": s_pid,
                "type": f"G{seed}P2",
                "group": seed,
            }

    print(
        f"Position map has {len(position_map)} players "
        f"({num_byes} bye slot(s) left empty)"
    )

    # Print initial pairings before conflict resolution
    print("\n=== INITIAL KO ROUND 1 PAIRINGS ===\n")
    print(
        f"{'Match':<5} {'Player A':<24} {'Verein A':<22} {' vs'} {'Player B':<24} {'Verein B':<25}"
    )
    print("-" * 100)
    matches = bracket_size // 2
    for i in range(matches):
        pos_a = 2 * i + 1
        pos_b = 2 * i + 2

        p_a_data = position_map.get(pos_a)
        p_b_data = position_map.get(pos_b)

        if not p_a_data and not p_b_data:
            continue  # phantom match, skip entirely

        if not p_a_data or not p_b_data:
            # One slot is a bye
            p_data = p_a_data or p_b_data
            p = players.get(p_data["pid"])
            if p:
                print(
                    f"{pos_a:<2}-{pos_b:<2} {p_data['type']:<5} {p['name']:<18} "
                    f"{p['verein']:<23} vs {'BYE'}"
                )
            continue

        p_a = players.get(p_a_data["pid"])
        p_b = players.get(p_b_data["pid"])

        if not p_a or not p_b:
            continue

        a_designation = p_a_data["type"]
        b_designation = p_b_data["type"]
        conflict = " ***" if p_a["verein"] == p_b["verein"] else ""
        print(
            f"{pos_a:<2}-{pos_b:<2} {a_designation:<5} {p_a['name']:<18} {p_a['verein']:<23} vs {b_designation:<5} {p_b['name']:<18} {p_b['verein']:<23}{conflict}"
        )

    # Check conflicts
    def check_conflicts(pos_map, bracket_size):
        conflicts = []
        matches = bracket_size // 2
        for i in range(matches):
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

    conflicts = check_conflicts(position_map, bracket_size)
    print(f"\nInitial club conflicts: {len(conflicts)}")
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

            new_conflicts = check_conflicts(position_map, bracket_size)
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

    # Read Phase 2 games early (for change detection in final pairings)
    spiele_raw = mdb_export(mdb_file, "tbl_Spiele")
    phase2_r1 = [
        r
        for r in spiele_raw
        if r.get("tsp_iPhase") == "2" and r.get("tsp_iRunde") == "1"
    ]

    # Print final pairings
    # Build a map of position -> old player ID from original Phase 2 games
    pos_to_old_player = {}
    for row in phase2_r1:
        try:
            pos_a = int(row.get("tsp_iPosAPlan", 0))
            pos_b = int(row.get("tsp_iPosBPlan", 0))
        except (ValueError, TypeError):
            continue
        old_a = row.get("tsp_refSpielerA_1")
        old_b = row.get("tsp_refSpielerB_1")
        if old_a:
            pos_to_old_player[pos_a] = old_a
        if old_b:
            pos_to_old_player[pos_b] = old_b

    # Red color for ANSI
    RED = "\033[91m"
    RESET = "\033[0m"

    print("\n=== FINAL KO ROUND 1 PAIRINGS ===\n")
    print(
        f"{'Match':<5} {'Player A':<24} {'Verein A':<22} {' vs'} {'Player B':<24} {'Verein B':<25}"
    )
    print("-" * 100)

    matches = bracket_size // 2
    for i in range(matches):
        pos_a = 2 * i + 1
        pos_b = 2 * i + 2

        p_a_data = position_map.get(pos_a)
        p_b_data = position_map.get(pos_b)

        if not p_a_data and not p_b_data:
            continue  # phantom match

        if not p_a_data or not p_b_data:
            # One slot is a bye: the present player advances without a match
            p_data = p_a_data or p_b_data
            p = players.get(p_data["pid"])
            if p:
                print(
                    f"{pos_a:<2}-{pos_b:<2} {p_data['type']:<5} {p['name']:<18} "
                    f"{p['verein']:<23} vs {'BYE':<5} (advances to round 2)"
                )
            continue

        p_a = players.get(p_a_data["pid"])
        p_b = players.get(p_b_data["pid"])

        if not p_a or not p_b:
            continue

        old_a = pos_to_old_player.get(pos_a)
        old_b = pos_to_old_player.get(pos_b)
        a_changed = old_a and old_a != p_a_data["pid"]
        b_changed = old_b and old_b != p_b_data["pid"]

        a_designation = p_a_data["type"]
        b_designation = p_b_data["type"]
        a_name = f"{RED}{p_a['name']}{RESET}" if a_changed else p_a["name"]
        b_name = f"{RED}{p_b['name']}{RESET}" if b_changed else p_b["name"]

        conflict = " ***" if p_a["verein"] == p_b["verein"] else ""
        print(
            f"{pos_a:<2}-{pos_b:<2} {a_designation:<5} {a_name:<18} {p_a['verein']:<23} vs {b_designation:<5} {b_name:<18} {p_b['verein']:<23}{conflict}"
        )
    print("-" * 95)

    if not phase2_r1:
        print("No Phase 2 Round 1 games found")
        return

    print(f"\nFound {len(phase2_r1)} Phase 2 Round 1 games")

    # Generate SQL updates.
    # Positions absent from position_map are byes: write -1 for that player slot.
    BYE_ID = "-1"
    updates = []
    for row in phase2_r1:
        try:
            pos_a = int(row.get("tsp_iPosAPlan", 0))
            pos_b = int(row.get("tsp_iPosBPlan", 0))
        except (ValueError, TypeError):
            continue

        new_a = position_map[pos_a]["pid"] if pos_a in position_map else BYE_ID
        new_b = position_map[pos_b]["pid"] if pos_b in position_map else BYE_ID

        # Skip fully phantom matches (no player on either side and both already -1)
        if new_a == BYE_ID and new_b == BYE_ID:
            continue

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
        print("\nFix the database...")
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

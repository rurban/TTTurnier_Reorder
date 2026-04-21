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


def generate_standard_seeding(size):
    """Generate standard tournament seeding for bracket size (power of 2)"""
    if size == 1:
        return [1]
    # Recursively build: take odd positions, then even positions reversed
    odd = generate_standard_seeding(size // 2)
    even = [x + size // 2 for x in odd]
    # Interleave: first half odds, second half evens
    result = []
    for i in range(len(odd)):
        result.append(odd[i])
        result.append(even[i])
    return result


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

    # Calculate bracket size (power of 2 >= number of advancers)
    num_groups = len(groups)
    num_advancers = num_groups * 2  # winner + second from each group

    # Find smallest power of 2 >= num_advancers
    bracket_size = 1
    while bracket_size < num_advancers:
        bracket_size *= 2

    num_byes = bracket_size - num_advancers

    if verbose:
        print(
            f"Groups: {num_groups}, Advancers: {num_advancers}, Bracket size: {bracket_size}, Byes: {num_byes}"
        )

    # Standard seeding for bracket size (power of 2)
    # This generates the standard tournament bracket seeding
    def generate_standard_seeding(size):
        """Generate standard tournament seeding for bracket size (power of 2)"""
        if size == 1:
            return [1]
        # Recursively build: take odd positions, then even positions reversed
        odd = generate_standard_seeding(size // 2)
        even = [x + size // 2 for x in odd]
        # Interleave: first half odds, second half evens
        result = []
        for i in range(len(odd)):
            result.append(odd[i])
            result.append(even[i])
        return result

    # Generate seeding: position -> seed (where seed 1 is best player)
    seeding = generate_standard_seeding(bracket_size)
    # Convert to position -> seed mapping
    pos_to_seed = {}
    for pos, seed in enumerate(seeding, start=1):
        pos_to_seed[pos] = seed

    # Now we need to map our groups to positions based on CLAUDE.md rules
    # But we only have num_groups groups, not necessarily 16
    # We'll adapt the CLAUDE.md pattern proportionally

    # Generate winner positions: alternate top/bottom halves (same as CLAUDE.md pattern)
    # For 32: 1,32,17,16,9,24,25,8,12,21,28,5,13,20,30,4
    n = num_groups * 2
    winner_positions = []
    top = list(range(1, n // 2 + 1))  # 1 to 16
    bottom = list(range(n, n // 2, -1))  # 32 to 17 (reversed)

    # Interleave: 1,32,17,16,9,24,25,8,...
    for i in range(len(top)):
        winner_positions.append(top[i])
        if i < len(bottom):
            winner_positions.append(bottom[i])
    winner_positions = winner_positions[:num_groups]

    # Generate second positions: in OTHER bracket half (NOT same match pair!)
    # If winner in upper half (1-16), second in lower half (17-32) at same relative position
    # If winner in lower half (17-32), second in upper half (1-16) at same relative position
    # This ensures winner vs second from SAME GROUP never meet in Round 1
    second_positions = []
    for pos in winner_positions:
        if pos <= n // 2:
            # Upper half -> lower half at same relative position
            second_positions.append(pos + n // 2)
        else:
            # Lower half -> upper half at same relative position
            second_positions.append(pos - n // 2)

    # Convert to dicts
    winner_pos = {i + 1: pos for i, pos in enumerate(winner_positions)}
    second_pos = {i + 1: pos for i, pos in enumerate(second_positions)}

    if verbose:
        print(f"Winner positions: {[winner_pos[i] for i in range(1, num_groups + 1)]}")
        print(f"Second positions: {[second_pos[i] for i in range(1, num_groups + 1)]}")

    # Build position map
    position_map = {}  # position -> {pid, type, group}
    for i, g in enumerate(groups, 1):
        pid = winners[i - 1] if i - 1 < len(winners) else "-1"
        if pid and pid != "-1" and pid in players:
            if i in winner_pos:
                position_map[winner_pos[i]] = {
                    "pid": pid,
                    "type": f"G{i}P1",
                    "group": i,
                }
            else:
                # Fallback: assign to next available position
                pass

        pid = seconds[i - 1] if i - 1 < len(seconds) else "-1"
        if pid and pid != "-1" and pid in players:
            if i in second_pos:
                position_map[second_pos[i]] = {
                    "pid": pid,
                    "type": f"G{i}P2",
                    "group": i,
                }
            else:
                # Fallback: assign to next available position
                pass

    print(f"Position map has {len(position_map)} players")

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

        if not p_a_data or not p_b_data:
            # Skip if one or both positions are empty (byes)
            continue

        p_a = players.get(p_a_data["pid"])
        p_b = players.get(p_b_data["pid"])

        if not p_a or not p_b:
            continue

        # Format: "1-2 G1P1 Name Verein - GxP2 Name Verein"
        a_designation = p_a_data["type"]  # e.g., "G1P1"
        b_designation = p_b_data["type"]  # e.g., "G2P2"

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

        if not p_a_data or not p_b_data:
            # Skip if one or both positions are empty (byes)
            continue

        p_a = players.get(p_a_data["pid"])
        p_b = players.get(p_b_data["pid"])

        if not p_a or not p_b:
            continue

        # Check if player was changed from original
        old_a = pos_to_old_player.get(pos_a)
        old_b = pos_to_old_player.get(pos_b)
        a_changed = old_a and old_a != p_a_data["pid"]
        b_changed = old_b and old_b != p_b_data["pid"]

        # Format: "1-2 G1P1 Name Verein - GxP2 Name Verein"
        a_designation = p_a_data["type"]  # e.g., "G1P1"
        b_designation = p_b_data["type"]  # e.g., "G2P2"

        # Add red color if changed
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

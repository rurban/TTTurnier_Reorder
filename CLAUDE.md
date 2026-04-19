TTTurnier Reorder
-----------------

Read the given Anmeldung_92067.xls excel (xml) spreadsheet.
Read tabs, the first is the general list of all "Turnieranmeldungen".
Then the list of tournament categories. Ie. "1 Damen B bis 1300(E)" and "2 Offen B 1300-1600(E)"

For each tab with more than 8 players, create a proper grouping.

- row 1: Turnieranmeldungen - Turnierklasse - Offen B 1300-1600 (Einzel) (Stand 16.04.2026 08:43:40)
- row 2: Nachname	Vorname	Geburtstag	Geschlecht	Verein	Verband-1	Verband-2	Verband-3	Intere ID	Verbands-SpielerNr	LivePZ	Lizenz	Ersatz
- Then all entries.

- Sort all players by column LivePZ (column k) downwards.
- For all players we need only Columns A: Namename, B: Vorname, E: Verein, K: LivePZ

Calculate the group sizes
-------------------------

With n=60 players we end up with 64 for the KO tree, under 32 it would be a KO tree of 32.
Divide n into m groups of 3-5 players, so that the two best players of each group end up in the KO tree.
E.g. with 60 we need either 16 groups of max 4. The first 4 of them with just 3 players.
Under 40 we get a division of 8 groups of 5 = 40.
Under 32 we get 8 * 4.

Sort the players
----------------
Sort in the players, sorted by LivePZ. (g1 = group, -1 first player, -2 second)

p1 => g1-1
p2 => g2-1
...
p16 => g16-1
p17 => g1-2
p18 => g1-2
...
p60 => g16-4

Re-arrange groups by Verein
----------------------------
Then re-arrange with special rules, so that no players of the same club are in the same group.
If there are more than m=16 players of a single Verein, that cannot be fulfilled. So there can be 2 of them in the same group.
When re-arranging them, check that the sum of distance of movements to the next group is minimal to keep the LivePZ differences minimal. I.e. when moving g2-2 to the next g3-2, g4-2, keep them in the very same position 2, so that the LivePZ differences are kept minimal.

Output the resulting groups in a single-file HTML, named after the inpir filename,
just with an HTML extension.
To be stored on the Windows desktop, same dir as the script/executable.
Color the re-arranged players red, the kept players black.

Re-arrange round 1 of the KO pairings  (Phase 2)
------------------------------------------------

The system does calculate the first round of the 2. Phase KO wrong.
We can create the pairings at first, close the program,
let a program adjust the mdb database, and read the turnier back in.

The pairings are at the `tbl_Spiele` table in the .mdb file.
You can export that via `mdb-export tbl_Spiele your.mdb >tbl_Spiele.csv`.

The games in Phase 2 start at row 98 and include:
- Round 1 (tsp_iRunde=1): Achtelfinale (last 16)
- Round 2: Viertelfinale (quarter finals)
- Round 3: Halbfinale (semi finals)
- etc.

Looking at the first Phase 2 game (line 98), it's:
35,72,2,0,951,0,1,962,0,31,97,,3,1,8,9,-7,7,,,,3,0,"04/19/26 11:58:00",,"04/19/26 12:26:00",167,,,,,,9898,1,,0

Where:
- tsp_refTurnier = 35
- tsp_refKlasse = 72
- tsp_iPhase = 2 (KO phase)
- tsp_refGruppe = 0 (no group)
- tsp_refSpielerA_1 = 951 (player ID)
- tsp_iPosAPlan = 1 (KO position)
- tsp_refSpielerB_1 = 962
- tsp_iPosBPlan = 31

Rules for the KO pairings (Phase 2)
-----------------------------------
The group winners of the first round are set at fixed positions.
With 16 winners, G1P1 is at position 1, G2P1 (2nd best) at 32, G3P1 at
17, G4P1 at 16.

For the phase 2, the LivePZ ranking should be ignored (unless this is a WTTV
Verband tournament). In the WTTV the LivePZ, i.e. the initial grpup
ranking, is more important than the actual group result. So G1P1 is from
the initial ranking of the group, not the actual group winner.

    1 => 1
    2 => 32
    3 => 17
    4 => 16
    5 => 9
    6 => 24
...
    15 => 30
    16 => 4

Now place 2nd places in the groups, such that:

1. The winner of that group is in the other major bracket (1-16 or 17-32),
   upper or lower.
2. 2nd places are distributed similarily as the winners accross the tree,
   just in the other bracket. G1P1 is up, G1P2 is down.
3. Duels of the same clubs should be avoided in the first KO round.

Language
--------

Do that in a programming language, usable on windows with a single executable. To be developed and tested on linux. For perl that would be perl2exe with perl-5.30.1 (https://www.indigostar.com/download/perl2exe-30.10-win.zip). For python `pyinstaller --onefile TTTurnier_Reorder.py`

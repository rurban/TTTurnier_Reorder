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


Do that in a programming language, usable on windows with a single executable. To be developed and tested on linux. For perl that would be perl2exe with perl-5.30.1 (https://www.indigostar.com/download/perl2exe-30.10-win.zip). For python `pyinstaller --onefile TTTurnier_Reorder.py`


#!/usr/bin/perl
# Copyright 2026 Reini Urban
# Licensed under GPL-3.0-or-later
# To be used with the Anmeldung*.xls files from the TTTurnier software
# for the first step of creating tournament groups.
# Written for the SVF Dresden, for the SEM-B and SEM-A tournaments in 2026.
# There is another step later for the KO-stage. See TTTurnier_KO_Reorder.pl

use strict;
use warnings;
use utf8;
use open ':std', ':encoding(UTF-8)';
use XML::LibXML;
use File::Basename qw(dirname basename);
use File::Spec;

# Spreadsheet XML namespace
my $NS = 'urn:schemas-microsoft-com:office:spreadsheet';

# ---------------------------------------------------------------------------
# Locate input XLS file
# ---------------------------------------------------------------------------
my ($xls_file, $out_file);

if (@ARGV) {
    $xls_file = $ARGV[0];
    $out_file  = $ARGV[1] if $ARGV[1];
}

unless ($xls_file && -f $xls_file) {
    # Auto-detect Anmeldung*.xls next to this script
    my $dir = dirname(File::Spec->rel2abs($0));
    my @cands = sort glob(File::Spec->catfile($dir, 'Anmeldung*.xls'));
    $xls_file = $cands[0] if @cands;
}

die "Usage: $0 <Anmeldung.xls> [output.html]\n"
    unless $xls_file && -f $xls_file;

# Default output: same directory, replace extension with _Gruppen.html
unless ($out_file) {
    my $base = $xls_file;
    $base =~ s/\.[^.\\\/]+$//;
    $out_file = "${base}_Gruppen.html";
}

# ---------------------------------------------------------------------------
# Parse XML / SpreadsheetML
# ---------------------------------------------------------------------------
my $parser = XML::LibXML->new();
my $doc    = $parser->parse_file($xls_file);
my $root   = $doc->documentElement();
my $xpc    = XML::LibXML::XPathContext->new($root);
$xpc->registerNs('s', $NS);

# Return array of cell values for a <s:Row>, respecting s:Index gaps
sub row_values {
    my ($row) = @_;
    my @out;
    my $cur = 0;
    for my $cell ($xpc->findnodes('s:Cell', $row)) {
        my $idx = $cell->getAttributeNS($NS, 'Index');
        $cur = $idx - 1 if defined $idx && $idx ne '';
        my ($data) = $xpc->findnodes('s:Data', $cell);
        $out[$cur] = $data ? $data->textContent() : '';
        $cur++;
    }
    return @out;
}

# ---------------------------------------------------------------------------
# Read worksheets
# ---------------------------------------------------------------------------
my @categories;

for my $ws ($xpc->findnodes('//s:Worksheet')) {
    my $sheet = $ws->getAttributeNS($NS, 'Name');
    next if $sheet eq 'Turnieranmeldungen';

    my @rows = $xpc->findnodes('s:Table/s:Row', $ws);
    next if @rows < 4;   # need title + header + at least 2 data rows

    # Row 0: title string
    my @t = row_values($rows[0]);
    my $title = $t[0] // $sheet;

    # Row 1: column headers
    my @hdrs = row_values($rows[1]);
    my %col;
    $col{$hdrs[$_]} = $_ for grep { defined $hdrs[$_] } 0..$#hdrs;

    my $ci_name   = $col{Nachname}  // 0;
    my $ci_vname  = $col{Vorname}   // 1;
    my $ci_verein = $col{Verein}    // 4;
    my $ci_pz     = $col{LivePZ}    // 10;

    # Rows 2+: players
    my @players;
    for my $i (2..$#rows) {
        my @v = row_values($rows[$i]);
        my $pz = $v[$ci_pz] // 0;
        $pz =~ s/[^\d]//g;
        push @players, {
            nachname => $v[$ci_name]   // '',
            vorname  => $v[$ci_vname]  // '',
            verein   => $v[$ci_verein] // '',
            livepz   => $pz + 0,
        };
    }

    next if @players <= 8;

    push @categories, {
        name    => $sheet,
        title   => $title,
        players => \@players,
    };
}

die "No tournament categories with more than 8 players found.\n"
    unless @categories;

# ---------------------------------------------------------------------------
# Build groups for each category
# ---------------------------------------------------------------------------

# Determine number of groups based on player count
#   n <= 32  : 8 groups of max 4
#   n <= 40  : 8 groups of max 5
#   n <= 64  : 16 groups of max 4
sub num_groups {
    my ($n) = @_;
    return 8  if $n <= 40;
    return 16 if $n <= 64;
    # Extend for larger fields: next power-of-2 / 4
    my $ko = 1;
    $ko *= 2 while $ko < $n;
    return $ko / 4;
}

for my $cat (@categories) {
    my @players = @{$cat->{players}};

    # Sort by LivePZ descending
    @players = sort { $b->{livepz} <=> $a->{livepz} } @players;
    $cat->{players} = \@players;

    my $n = scalar @players;
    my $m = num_groups($n);

    # Initial seeding: round-robin by rank
    # p[0]->g0, p[1]->g1, ..., p[m-1]->g(m-1), p[m]->g0, ...
    my @groups;
    for my $i (0..$#players) {
        push @{$groups[$i % $m]}, { %{$players[$i]}, orig_rank => $i };
    }

    # -------------------------------------------------------------------
    # Re-arrange: no two players of the same Verein in the same group.
    # Swap conflicts with the partner that minimises |LivePZ difference|.
    # We allow a conflict to remain if no valid swap exists (i.e. more
    # than m players from a single Verein).
    # -------------------------------------------------------------------
    my $changed = 1;
    my $iters   = 0;
    my $max_iters = $m * $n + 10;   # generous upper bound

    while ($changed && $iters++ < $max_iters) {
        $changed = 0;

        for my $gi (0..$m-1) {
            # Count Verein occurrences in this group
            my %vc;
            $vc{$_->{verein}}++ for @{$groups[$gi]};

            for my $verein (sort keys %vc) {
                next if $vc{$verein} < 2;

                # Player to move: last (weakest) duplicate in the group
                my ($move_pi) = reverse grep {
                    $groups[$gi][$_]{verein} eq $verein
                } 0..$#{$groups[$gi]};
                my $mp = $groups[$gi][$move_pi];

                my ($best_gj, $best_pj, $best_cost) = (-1, -1, 1e18);

                for my $gj (0..$m-1) {
                    next if $gj == $gi;

                    for my $pj (0..$#{$groups[$gj]}) {
                        my $cp = $groups[$gj][$pj];
                        next if $cp->{verein} eq $verein; # keeps conflict in gi

                        # Simulate swap and check both groups for new conflicts
                        my %v_gi;
                        for my $k (0..$#{$groups[$gi]}) {
                            my $p = ($k == $move_pi) ? $cp : $groups[$gi][$k];
                            $v_gi{$p->{verein}}++;
                        }
                        next if ($v_gi{$cp->{verein}} // 0) > 1;

                        my %v_gj;
                        for my $k (0..$#{$groups[$gj]}) {
                            my $p = ($k == $pj) ? $mp : $groups[$gj][$k];
                            $v_gj{$p->{verein}}++;
                        }
                        next if ($v_gj{$mp->{verein}} // 0) > 1;

                        my $cost = abs($mp->{livepz} - $cp->{livepz});
                        if ($cost < $best_cost) {
                            $best_cost = $cost;
                            $best_gj   = $gj;
                            $best_pj   = $pj;
                        }
                    }
                }

                if ($best_gj >= 0) {
                    ( $groups[$gi][$move_pi], $groups[$best_gj][$best_pj] ) =
                    ( $groups[$best_gj][$best_pj], $groups[$gi][$move_pi] );
                    $changed = 1;
                }
            }
        }
    }

    # Mark moved players (final group != originally seeded group)
    for my $gi (0..$m-1) {
        for my $p (@{$groups[$gi]}) {
            $p->{moved} = ( ($p->{orig_rank} % $m) != $gi ) ? 1 : 0;
        }
    }

    $cat->{groups} = \@groups;
    $cat->{m}      = $m;
}

# ---------------------------------------------------------------------------
# HTML output
# ---------------------------------------------------------------------------
open(my $fh, '>:encoding(UTF-8)', $out_file)
    or die "Cannot write '$out_file': $!\n";

sub esc {
    my $s = shift // '';
    $s =~ s/&/&amp;/g;
    $s =~ s/</&lt;/g;
    $s =~ s/>/&gt;/g;
    $s =~ s/"/&quot;/g;
    return $s;
}

print $fh <<'HTML';
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Turniergruppen</title>
<style>
  body   { font-family: Arial, sans-serif; font-size: 12px; margin: 16px; background: #fff; color: #000; }
  h1     { font-size: 15px; margin: 0 0 4px 0; padding-bottom: 4px; border-bottom: 2px solid #444; }
  .meta  { font-size: 11px; color: #666; margin: 0 0 12px 0; }
  .wrap  { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 32px; page-break-after: always; }
  .grp   { border: 1px solid #aaa; border-radius: 3px; min-width: 200px; }
  .ghdr  { background: #d4d4d4; font-weight: bold; text-align: center;
           padding: 3px 6px; font-size: 12px; }
  table  { border-collapse: collapse; width: 100%; }
  td     { padding: 2px 6px; border-top: 1px solid #e8e8e8; white-space: nowrap; }
  .nr    { color: #aaa; text-align: right; width: 18px; font-size: 10px; }
  .pz    { text-align: right; color: #555; width: 36px; }
  .red   { color: red; }
  .blk   { color: black; }
  @media print { .wrap { page-break-after: always; } }
</style>
</head>
<body>
HTML

for my $cat (@categories) {
    my @groups  = @{$cat->{groups}};
    my $m       = $cat->{m};
    my $n       = scalar @{$cat->{players}};
    my $moved_n = 0;
    for my $gi (0..$#groups) {
        $moved_n += grep { $_->{moved} } @{$groups[$gi]};
    }

    print $fh '<h1>', esc($cat->{title}), "</h1>\n";
    printf $fh "<p class=\"meta\">%d Spieler &bull; %d Gruppen%s</p>\n",
        $n, $m,
        ($moved_n ? " &bull; <span class=\"red\">$moved_n umgesetzt</span>" : '');

    print $fh "<div class=\"wrap\">\n";

    for my $gi (0..$#groups) {
        # Sort within group by LivePZ descending for display
        my @grp = sort { $b->{livepz} <=> $a->{livepz} } @{$groups[$gi]};

        printf $fh "<div class=\"grp\"><div class=\"ghdr\">Gruppe %d</div>\n", $gi + 1;
        print $fh "<table>\n";

        my $rank = 1;
        for my $p (@grp) {
            my $cls = $p->{moved} ? 'red' : 'blk';
            printf $fh
                "<tr><td class=\"nr\">%d</td>"
                . "<td class=\"%s\">%s</td>"
                . "<td class=\"%s\">%s</td>"
                . "<td class=\"pz %s\">%s</td></tr>\n",
                $rank++,
                $cls, esc("$p->{nachname}, $p->{vorname}"),
                $cls, esc($p->{verein}),
                $cls, ($p->{livepz} || '');
        }

        print $fh "</table></div>\n";
    }

    print $fh "</div>\n";
}

print $fh "</body>\n</html>\n";
close $fh;

printf "Written: %s\n", $out_file;
printf "Categories processed: %d\n", scalar @categories;
for my $cat (@categories) {
    my $moved = 0;
    $moved += grep { $_->{moved} } @$_ for @{$cat->{groups}};
    printf "  %-30s  %2d players  %2d groups  %d moved\n",
        $cat->{name}, scalar(@{$cat->{players}}), $cat->{m}, $moved;
}

__END__

=head1 NAME

TTTurnier_Reorder - Group draw generator for TTTurnier tournaments

=head1 SYNOPSIS

    perl TTTurnier_Reorder.pl [<Anmeldung.xls> [<output.html>]]

    # Auto-detect Anmeldung*.xls in the script directory, write next to it:
    perl TTTurnier_Reorder.pl

    # Explicit input, auto output (Anmeldung_92067_Gruppen.html):
    perl TTTurnier_Reorder.pl Anmeldung_92067.xls

    # Explicit input and output:
    perl TTTurnier_Reorder.pl Anmeldung_92067.xls Gruppen.html

    # On Windows (compiled with PAR::Packer):
    TTTurnier_Reorder.exe Anmeldung_92067.xls

=head1 DESCRIPTION

Reads a TTTurnier registration export in SpreadsheetML format (C<.xls>).
The file contains one overview sheet (C<Turnieranmeldungen>) and one sheet
per tournament category.  Each category sheet with more than 8 players is
processed into groups.

Written for the SVF Dresden, for the SEM-B and SEM-A tournaments in 2026.
There is another step later for the KO-stage. See F<TTTurnier_KO_Reorder.pl>

=head2 Algorithm

=over 4

=item 1. B<Sort> all players by LivePZ (rating) descending.

=item 2. B<Determine group count> based on field size:

    n <= 40  =>  8 groups  (max 4-5 players each)
    n <= 64  => 16 groups  (max 4 players each)
    larger   =>  next power-of-2 / 4 groups

=item 3. B<Initial seeding> by round-robin:
p1->G1, p2->G2, ..., pm->Gm, p(m+1)->G1, ...
This keeps the LivePZ spread within each group as even as possible.

=item 4. B<Re-arrange by Verein (club)>: iteratively swap players to
ensure no two players from the same club share a group.  Each swap
chooses the partner from another group that minimises the absolute
LivePZ difference, so the overall rating balance is disturbed as little
as possible.  If a club has more players than there are groups, some
sharing is unavoidable and the remaining conflict is left in place.

=back

=head2 Output

A self-contained HTML file with one card per group.  Players who were
moved during the re-arrangement step are coloured B<red>; players who
stayed in their originally seeded group are shown in black.

The output file is written next to the input C<.xls> file (or to
C<ARGV[1]> if supplied), with the suffix C<_Gruppen.html>.

=head1 DEPENDENCIES

L<XML::LibXML> (included in Strawberry Perl for Windows).

To build a standalone Windows executable:

    pp -o TTTurnier_Reorder.exe TTTurnier_Reorder.pl
or
    perl2exe TTTurnier_Reorder.pl

=head1 AUTHOR

Reini Urban

=head1 LICENSE

GPL-3.0-or-later

=cut

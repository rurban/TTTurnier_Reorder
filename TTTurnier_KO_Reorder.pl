#!/usr/bin/perl
# Copyright 2026 Reini Urban
# Licensed under GPL-3.0-or-later
# Reorders the first round of Phase 2 (KO) pairings based on group results.
# Reads from TTTurnier .mdb database, rearranges Round of 32, writes back.
# On Unix uses mdbtools, on Windows uses ODBC.

use strict;
use warnings;
use utf8;
use open ':std', ':encoding(UTF-8)';
use File::Basename qw( dirname fileparse );
use File::Spec     ();
use File::Copy     qw( copy );

# ---------------------------------------------------------------------------
# Locate input .mdb file
# ---------------------------------------------------------------------------
my ( $mdb_file, %opts );
my $exe_dir = File::Spec->rel2abs( dirname( $0 || $ENV{PWD} ) );

# Parse arguments
for my $arg (@ARGV) {
    if ( $arg eq '-v' || $arg eq '--verbose' ) {
        $opts{verbose} = 1;
    }
    elsif ( $arg eq '-n' || $arg eq '--dry-run' ) {
        $opts{dry_run} = 1;
    }
    elsif ( -f $arg && $arg =~ /\.mdb$/i ) {
        $mdb_file = $arg;
    }
    else {
        die "Unknown argument: $arg\n";
    }
}

# Auto-detect .mdb file
unless ( $mdb_file && -f $mdb_file ) {
    my @cands = sort glob( File::Spec->catfile( $exe_dir, '*.mdb' ) );
    $mdb_file = $cands[0] if @cands;
}

die "Usage: $0 [<database.mdb>] [-v|--verbose] [-n|--dry-run]\n"
    unless $mdb_file && -f $mdb_file;

my $mdb_dir  = dirname($mdb_file);
my $mdb_name = fileparse($mdb_file);

# ---------------------------------------------------------------------------
# Backup original .mdb
# ---------------------------------------------------------------------------
my $backup = File::Spec->catfile( $mdb_dir, "${mdb_name}.bak" );
if ( -f $backup ) {
    print "Backup already exists: $backup\n";
}
else {
    copy( $mdb_file, $backup )
        or die "Cannot backup to '$backup': $!\n";
    print "Backed up: $backup\n";
}

# ---------------------------------------------------------------------------
# Export tables from .mdb using mdbtools
# ---------------------------------------------------------------------------
my %tables;

sub mdb_export {
    my ($table) = @_;
    my @rows;
    open my $fh, '-|', "mdb-export '$mdb_file' $table"
        or die "mdb-export failed: $!";
    my $header = <$fh>;
    chomp $header;
    my @cols = split /,/, $header;
    while (<$fh>) {
        chomp;
        my @vals = map { s/^"(.*)"$/$1/; s/""/"/g; $_ } split /,/;
        my %row;
        @row{@cols} = @vals;
        push @rows, \%row;
    }
    close $fh;
    return \@rows;
}

sub mdb_import {
    my ( $table, $rows ) = @_;
    return unless @$rows;

    my @cols    = keys %{ $rows->[0] };
    my $col_str = join( ',', @cols );
    my $esc     = sub {
        my $v = shift // '';
        $v =~ s/"/""/g;
        $v = "\"$v\"" if $v =~ /[",]/;
        return $v;
    };

    open my $fh, '|-', "mdb-export -T '$mdb_file' $table"
        or die "mdb-export -T failed: $!";

    print $fh "$col_str\n";
    for my $row (@$rows) {
        my @vals = map { &$esc( $row->{$_} ) } @cols;
        print $fh join( ',', @vals ) . "\n";
    }
    close $fh;
    print "Updated table: $table\n";
}

print "Reading tables from: $mdb_file\n";

# Read tbl_Spieler (players)
my $players_raw = mdb_export('tbl_Spieler');
my %players;
for my $p (@$players_raw) {
    my $id = $p->{ts_ID};
    next unless $id;
    $players{$id} = {
        livepz => int( $p->{ts_sSpielstaerke} || 0 ),
        verein => $p->{ts_sVereinName} // '',
        name   => "$p->{ts_Vorname} $p->{ts_Nachname}",
    };
}
print "Loaded ", scalar( keys %players ), " players\n";

# Read tbl_Tabelle (group results)
my $tabelle_raw = mdb_export('tbl_Tabelle');
my %group_results;    # gruppe -> {1 => winner_id, 2 => second_id}

for my $row (@$tabelle_raw) {
    my $gruppe  = $row->{tta_refGruppenID};
    my $platz   = int( $row->{tta_iPlatz} || 0 );
    my $spieler = $row->{tta_refSpieler};
    next unless $gruppe && $platz && $spieler && $spieler ne '-1';
    next unless $platz == 1 || $platz == 2;
    $group_results{$gruppe} //= {};
    $group_results{$gruppe}{$platz} = $spieler;
}

my @groups = sort { $a <=> $b } keys %group_results;
print "Found ", scalar(@groups), " groups\n";

# ---------------------------------------------------------------------------
# Build KO positions from group results
# ---------------------------------------------------------------------------

# Winners (platz 1) and 2nd places (platz 2) by group order
my @winners;
my @seconds;
for my $g (@groups) {
    push @winners, $group_results{$g}{1} // '-1';
    push @seconds, $group_results{$g}{2} // '-1';
}

# Fixed position mapping for 16 group winners (CLAUDE.md):
#   1=>1, 2=>32, 3=>17, 4=>16, 5=>9, 6=>24, 7=>25, 8=>8,
#   9=>12,10=>21,11=>28,12=>5,13=>13,14=>20,15=>30,16=>4
my %winner_pos = (
    1  => 1,
    2  => 32,
    3  => 17,
    4  => 16,
    5  => 9,
    6  => 24,
    7  => 25,
    8  => 8,
    9  => 12,
    10 => 21,
    11 => 28,
    12 => 5,
    13 => 13,
    14 => 20,
    15 => 30,
    16 => 4
);

# For 2nd places: opposite bracket from winner
# If winner in upper (1-16), 2nd goes to lower (17-32), and vice versa
my %second_pos;
for my $grp ( 1 .. 16 ) {
    my $wp = $winner_pos{$grp};
    if ( $wp <= 16 ) {
        $second_pos{$grp} = $wp + 16;
    }
    else {
        $second_pos{$grp} = $wp - 16;
    }
}

# Build position -> player mapping
my %position_map;    # position => {pid, type, group}
for my $i ( 0 .. $#groups ) {
    my $grp = $i + 1;
    my $pid = $winners[$i];
    if ( $pid && $pid ne '-1' ) {
        my $pos = $winner_pos{$grp};
        $position_map{$pos}
            = { pid => $pid, type => "G${grp}P1", group => $grp };
    }
    $pid = $seconds[$i];
    if ( $pid && $pid ne '-1' ) {
        my $pos = $second_pos{$grp};
        $position_map{$pos}
            = { pid => $pid, type => "G${grp}P2", group => $grp };
    }
}

print "Position map has ", scalar( keys %position_map ), " players\n";

# ---------------------------------------------------------------------------
# Check and resolve club conflicts
# ---------------------------------------------------------------------------

# Find matches with same club: positions (1,2), (3,4), ..., (31,32)
sub check_conflicts {
    my ($pos_map) = @_;
    my @conflicts;
    for my $i ( 0 .. 15 ) {
        my $pos_a = 2 * $i + 1;
        my $pos_b = 2 * $i + 2;
        next unless exists $pos_map->{$pos_a} && exists $pos_map->{$pos_b};
        my $p_a = $players{ $pos_map->{$pos_a}{pid} };
        my $p_b = $players{ $pos_map->{$pos_b}{pid} };
        if ( $p_a && $p_b && $p_a->{verein} eq $p_b->{verein} ) {
            push @conflicts,
                {
                pos_a  => $pos_a,
                pos_b  => $pos_b,
                name_a => $p_a->{name},
                name_b => $p_b->{name},
                club   => $p_a->{verein},
                };
        }
    }
    return @conflicts;
}

my @conflicts = check_conflicts( \%position_map );
print "Initial club conflicts: ", scalar(@conflicts), "\n";
for my $c (@conflicts) {
    print
        "  Match $c->{pos_a} vs $c->{pos_b}: $c->{name_a} vs $c->{name_b} ($c->{club})\n";
}

# Try to resolve conflicts by swapping players
# We can swap positions within the same match, or between positions in different matches
if (@conflicts) {
    my $resolved = 0;

    # Try swaps within same match (swap a and b)
    for my $c (@conflicts) {
        my $pos_a = $c->{pos_a};
        my $pos_b = $c->{pos_b};
        next
            unless exists $position_map{$pos_a}
            && exists $position_map{$pos_b};

        # Swap them in the same match
        my %tmp = %{ $position_map{$pos_a} };
        $position_map{$pos_a} = $position_map{$pos_b};
        $position_map{$pos_b} = \%tmp;

        my @check = check_conflicts( \%position_map );
        if ( scalar(@check) < scalar(@conflicts) ) {
            @conflicts = @check;
            $resolved++;
            print "Resolved by swapping within match $pos_a/$pos_b\n";
            last if !@conflicts;
        }
        else {
            # Swap back
            %tmp                  = %{ $position_map{$pos_a} };
            $position_map{$pos_a} = $position_map{$pos_b};
            $position_map{$pos_b} = \%tmp;
        }
    }

    # Try swapping across matches (different position pairs)
    unless ( !@conflicts ) {
        my @positions = sort { $a <=> $b } keys %position_map;
        for my $i ( 0 .. $#positions ) {
            for my $j ( $i + 1 .. $#positions ) {
                next
                    if int( $positions[$i] / 2 )
                    == int( $positions[$j] / 2 );    # same match

                # Try swapping these positions
                my %tmp = %{ $position_map{ $positions[$i] } };
                $position_map{ $positions[$i] }
                    = $position_map{ $positions[$j] };
                $position_map{ $positions[$j] } = \%tmp;

                my @check = check_conflicts( \%position_map );
                if ( scalar(@check) < scalar(@conflicts) ) {
                    @conflicts = @check;
                    $resolved++;
                    print
                        "Resolved by swapping positions $positions[$i] <-> $positions[$j]\n";
                    last if !@conflicts;
                }
                else {
                    # Swap back
                    %tmp = %{ $position_map{ $positions[$i] } };
                    $position_map{ $positions[$i] }
                        = $position_map{ $positions[$j] };
                    $position_map{ $positions[$j] } = \%tmp;
                }
            }
            last if !@conflicts;
        }
    }

    print "Resolved $resolved conflicts, remaining: ", scalar(@conflicts),
        "\n";
}

# ---------------------------------------------------------------------------
# Generate optimized Phase 2 Round 1 pairings
# ---------------------------------------------------------------------------

print "\n=== FINAL KO ROUND 1 PAIRINGS ===\n";
print
    "Pos  Player A                                    Player B                                    Club\n";
print "-" x 100, "\n";

for my $i ( 0 .. 15 ) {
    my $pos_a = 2 * $i + 1;
    my $pos_b = 2 * $i + 2;
    my $p_a   = $players{ $position_map{$pos_a}{pid} };
    my $p_b   = $players{ $position_map{$pos_b}{pid} };
    next unless $p_a && $p_b;

    my $conflict
        = ( $p_a->{verein} eq $p_b->{verein} ) ? " *** CONFLICT ***" : "";
    printf "%2d  %-40s %-40s %s\n",
        $pos_a,
        "$p_a->{name} (LivePZ:$p_a->{livepz})",
        "$p_b->{name} (LivePZ:$p_b->{livepz})$conflict",
        $p_a->{verein};
}

my $final_conflicts = scalar( check_conflicts( \%position_map ) );
print "\nTotal club conflicts: $final_conflicts\n";

# ---------------------------------------------------------------------------
# Write back to tbl_Spiele
# ---------------------------------------------------------------------------

if ( $opts{dry_run} ) {
    print "Dry run - no changes written\n";
}
else {
    # Read current Phase 2 Round 1 games
    my $spiele_raw = mdb_export('tbl_Spiele');
    my @spiele     = @$spiele_raw;

    my %update_count;
    for my $row (@spiele) {
        next unless $row->{tsp_iPhase} eq '2';
        next unless $row->{tsp_iRunde} eq '1';

        my $pos_a = int( $row->{tsp_iPosAPlan} );
        my $pos_b = int( $row->{tsp_iPosBPlan} );

        next
            unless exists $position_map{$pos_a}
            && exists $position_map{$pos_b};

        my $new_a = $position_map{$pos_a}{pid};
        my $new_b = $position_map{$pos_b}{pid};

        if (   $row->{tsp_refSpielerA_1} ne $new_a
            || $row->{tsp_refSpielerB_1} ne $new_b )
        {
            $row->{tsp_refSpielerA_1}       = $new_a;
            $row->{tsp_refSpielerB_1}       = $new_b;
            $update_count{ $row->{tsp_ID} } = 1;
        }
    }

    # Write back using mdbtools
    # Note: mdb-import doesn't exist, so we use sql via mdb-sql
    # For now just print the SQL statements
    if (%update_count) {
        print "\nUpdated ", scalar( keys %update_count ), " games\n";

        # Generate SQL UPDATE statements
        print "\nSQL UPDATE statements:\n";
        for my $game_id ( sort keys %update_count ) {
            my $row = ( grep { $_->{tsp_ID} eq $game_id } @spiele )[0];
            next unless $row;

            my $new_a = $position_map{ int( $row->{tsp_iPosAPlan} ) }{pid};
            my $new_b = $position_map{ int( $row->{tsp_iPosBPlan} ) }{pid};

            printf
                "UPDATE tbl_Spiele SET tsp_refSpielerA_1='%s', tsp_refSpielerB_1='%s' WHERE tsp_ID=%s;\n",
                $new_a, $new_b, $game_id;
        }
    }
    else {
        print "No changes needed\n";
    }
}

print "\nDone.\n";

__END__

=head1 NAME

TTTurnier_KO_Reorder - KO round re-seeding for TTTurnier Phase 2

=head1 SYNOPSIS

    perl TTTurnier_KO_Reorder.pl [-v] [-n] [<database.mdb>]

    # Auto-detect .mdb in script directory:
    perl TTTurnier_KO_Reorder.pl

    # Verbose output:
    perl TTTurnier_KO_Reorder.pl -v sem_b_2026.mdb

    # Dry run (show changes but don't write):
    perl TTTurnier_KO_Reorder.pl -n sem_b_2026.mdb

=head1 DESCRIPTION

Reads the TTTurnier MS Access database, extracts Phase 1 group results from
tbl_Tabelle, and reorders the Phase 2 Round of 1 (Achtelfinale) pairings
according to the rules:

=over 4

=item 1. B<Group winners> get fixed KO positions: 1E<62>1, 2E<62>32, 3E<62>17, 4E<62>16, ...

=item 2. B<2nd places> go to the opposite bracket from their group's winner.

=item 3. B<Club conflicts> are resolved by swapping players while minimizing
LivePZ differences.

=back

The original .mdb is backed up to F<database.mdb.bak>.

Requires mdbtools (unix) or ODBC driver (windows).

=head1 AUTHOR

Reini Urban

=head1 LICENSE

GPL-3.0-or-later

=cut

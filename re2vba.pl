#!perl
# Convert a Perl regex to a VBA implementation somewhat like that regex.
# Copyright (C) Chris White 2018.  Licensed MIT.
#
use 5.018;
use strict;
use warnings;
use Data::Dumper;
use Carp;
use Getopt::Long;
use Pod::Usage;

use constant { true => !!1, false => !!0 };
use constant EXIT_OK        => 0;   # success
use constant EXIT_PROC_ERR  => 1;   # error during processing
use constant EXIT_PARAM_ERR => 2;   # couldn't understand the command line

exit Main();

# Test and save a piece
sub stash_piece
{
    my ($hrPieces, $piecename, $piecetext) = @_;

    eval { my $re = qr{$piecetext} };
    croak "$piecename is not a valid regex: $@" if $@;

    # We don't support all group types
    my @bad_groups = $piecetext =~ m{
        (?<!\\)     # escaped \( are OK
        \(\?        # begin a special group
        [^:<]       # we only support non-capturing (?:) and named (?<>)
    }gx;
    croak "Unsupported groups: ", join(' ', @bad_groups) if @bad_groups;

    $hrPieces->{$piecename} = $piecetext;
} #stash_piece()

sub Main
{
    # Args
    my %opts = (dim=>true, private=>false, quiet=>false, indent=>true);
    GetOptions(\%opts,
        'usage|?', 'help|h', 'man',     # options we handle here
        "dim!",     # whether to print declarations
        "private",  # if true, use Private instead of Dim.  Ignored if --nodim
        "quiet|q",
        "indent!",  # if true, indent the output.
    )
    or pod2usage(-verbose => 0, -exitval => EXIT_PARAM_ERR);    # unknown opt

    # Help, if requested
    pod2usage(-verbose => 0, -exitval => EXIT_PROC_ERR) if $opts{usage};
    pod2usage(-verbose => 1, -exitval => EXIT_PROC_ERR) if $opts{help};
    pod2usage(-verbose => 2, -exitval => EXIT_PROC_ERR) if $opts{man};

    # Main input loop
    my %pieces;
    my $piecename="";
    my $piecetext="";
    my $mainpiece="";

    while(<>) {
        chomp;
        next if /^\s*#/;
        last if /^__END__\b/;

        s{\(\?#[^\)]*\)}{}g;        # remove comment groups
        s{\s+$}{};                  # remove trailing whitespace
        my @fields = split(/\s+/, $_, 2);
        say STDERR "${.}: ", join(':',@fields,'') unless $opts{quiet};

        if($fields[0] && $fields[1]) {    # a new piece
            if($piecename && $piecetext) {  # Finish the last piece
                stash_piece \%pieces, $piecename, $piecetext;
                $mainpiece = $piecename unless $mainpiece;
            }
            $piecename = $fields[0];    # start the new piece
            $piecetext = $fields[1];
        } elsif($fields[1]) {           # empty fields[0] => continuation
            $piecetext .= $fields[1];
        }
    } #main input loop

    # Finish the last piece, if any
    if($piecename && $piecetext) {
        stash_piece \%pieces, $piecename, $piecetext;
        $mainpiece = $piecename unless $mainpiece;
    }

    say STDERR "Found pieces: ", Dumper(\%pieces) unless $opts{quiet};

    die "No main piece found" unless $mainpiece && $pieces{$mainpiece};

    # Assemble the pieces into one regex
    my $full_regex = $pieces{$mainpiece};
    while($full_regex =~ s{\(\?<=([^\)]+)\)}{$pieces{$1}}gx) {
        die "Unknown piece $1" unless $pieces{$1};
    }

    say STDERR "Full regex is -$full_regex-" unless $opts{quiet};

    # Disallow \\( so I don't have to count backslashes to see if it's
    # even or odd.
    if($full_regex =~ m{\\\\\(}) {
        # Mark the error location
        say STDERR   "Full regex is -$full_regex-" if $opts{quiet};
        my $spacer = '               ' . (' ' x $-[0]);

        say STDERR $spacer, '^';
        die 'Unfortunately, I can\'t handle `\\\\(` (you can insert `.{0,0}`' .
            ' as a spacer if necessary)';
    }

    # Find named- or non-capturing groups in the regexes.
    my %names;
    my $groupidx=0;

    while( $full_regex =~ m{
        (?<!\\)     # Ignore escaped parens.
            # NOTE: this fails for `\\(foo)`, which should not be ignored.
            # TODO see https://stackoverflow.com/q/9613522/2877364
        \(          # open a group
        (?|
            (\?:)   # It's a non-capturing group
            |
            (\?<([^>]+)>)     # It's a named capturing group
        )?
    }gx) {
        my ($match_start, $match_end) = ($-[0], $+[0]);
        my $pos = pos $full_regex;

        my ($type_start, $type_end) = ($-[1], $+[1]);
        my $group_type = $1;    # may be undef
        my ($name_start, $name_end) = ($-[2], $+[2]);
        my $group_name = $2;    # may be undef

        if($group_type) {
            # Remove the group type, since VBScript can't handle those
            substr($full_regex, $type_start, $type_end-$type_start) = '';
            pos($full_regex) = $pos - ($type_end-$type_start);

            # Stash offset for named groups.  Note: for multiple occurrences
            # of a group name, only the last will be preserved.
            $names{$groupidx} = $group_name if $group_name;
        }
        ++$groupidx;
    } # for each group

    # Escape the double-quotes for VBA
    $full_regex =~ s{"}{""}g;

    # Process the definitions, and print them if desired
    my $I = ($opts{indent} ? ' ' : '');     # indent string
    my $I4 = $I x 4;
    my $I8 = $I x 8;

    say <<"EOT";
$I4' The following code is from the output of \`re2vba.pl vim-regex.txt\`.
$I4' DO NOT MODIFY HERE.  If you need to change it, modify vim-regex.txt
$I4' and re-run re2vba.pl.
EOT

    for(my $idx=0; $idx < $groupidx; ++$idx) {
        next unless exists $names{$idx};
        my $name = $names{$idx};
        $name = uc $name;
        $name =~ s{[^a-zA-Z0-9]}{_}g;
        $names{$idx} = $name;
        say(($opts{private} ? 'Private ' : ($I4) . 'Dim '),
            "RESM_$name As Long") if $opts{dim};
    }
    say(($opts{private} ? 'Private ' : ($I4) . 'Dim '),
        "RE_PAT As String\n") if $opts{dim};

    # Print the regex, with lines broken
    say $I4, "RE_PAT = _";
    while($full_regex =~ m{(.{0,60})}g) {
        say $I8, "\"$1\" & _" if $1;
    }
    say $I8, "\"\"";

    # Print the submatch numbers
    for(my $idx=0; $idx < $groupidx; ++$idx) {
        say $I4, "RESM_$names{$idx} = $idx" if exists $names{$idx};
    }

    say "\n${I4}' End of generated code";

    return 0;
} #Main()

__END__

=pod

=head1 NAME

re2vba.pl - Convert a Perl regex to a VBA implementation somewhat like that regex.

=head1 USAGE

    re2vba.pl [-options] [input files (stdin if none is given)]

=head1 OPTIONS

=over

=item --nodim

If given on the command line, do not print the C<Dim> statements.

=item --private

If given, use C<Private> instead of C<Dim>.

=item -q, --quiet

Do not print the diagnostic messages while running.

=item --noindent

If given on the command line, do not indent the output.

=back

=head1 INPUT FORMAT

    # comment to eol (hash must be first non-whitespace on the line)
    <piece name> <regex text>
    [<ws> <regex text continued>]
    __END__ (end of file)

Comment groups (C<(?#...)>) can be used, but must not span lines.
Whitespace remaining at the end of any line after removing comment groups
is ignored.
The first piece given is the main one.

In each piece of regex text, C<< (?<name>) >> defines a group that will be captured
and given a submatch number.  Backreferences are not currently processed.
C<< (?<=piecename) >> is replaced with the text of piece C<piecename>.
(In a real regex, that would be a positive lookbehind assertion, but VBScript
doesn't support those, so we can repurpose it.)
Each piece can be referenced only once.

=head2 Example input

    # piece3 is the main regex
    piece3  (?<=piece1)text here to make it long(?<=piece2)(?:foo)?(bar)?
            (?<lastone>9+)
    piece2  ((?<upperalpha>[A-Z])(?<p2>something))\\a+
            "[0-9]"
    piece1  [a-z](?<p1>thing)  (?#comment!)
    __END__

=head1 OUTPUT FORMAT

The tool outputs diagnostics on STDERR (unless --quite) and VBA source on
STDOUT.  The VBA source defines C<RE_PAT>, which is a string of the regex
pattern.  It also defines C<RESM_*> variables as C<Long>.  Those are the
submatch numbers of the various named groups in the input.  Each named group is
uppercased, and all non-letter/non-digit characters are replaced with
underscores.

=head2 Example output

    Dim RE_PAT As String
    Dim RESM_P1 As Long
    Dim RESM_UPPERALPHA As Long
    Dim RESM_P2 As Long
    Dim RESM_LASTONE As Long

    RE_PAT = _
        "[a-z](thing)text here to make it long(([A-Z])(something))\\a" & _
        "+""[0-9]""(foo)?(bar)?(9+)" & _
        ""
    RESM_P1 = 0
    RESM_UPPERALPHA = 2
    RESM_P2 = 3
    RESM_LASTONE = 6

=head1 COPYRIGHT

Copyright (C) Chris White 2018.  Licensed Artistic 2.0.

=cut

# vi: set ts=4 sts=4 sw=4 et ai ff=unix: #

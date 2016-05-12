use strict;
use lib 'lib/';
use Carp;
use Data::Dumper;
use Eldhelm::Util::CommandLine;
use Eldhelm::Util::FileSystem;

my %args;

sub nfo($) {
	my ($str) = @_;
	print "$str\n";
}

sub nag($) {
	my ($str) = @_;
	warn $str;
}

sub checkDeps() {
	my @deps = qw(Spreadsheet::ParseExcel Spreadsheet::XLSX);
	my @install;
	foreach my $d (@deps) {
		eval "use $d; 1" or do {
			nfo "I need $d";
			push @install, $d;
		};
	}
	return unless @install;

	nfo 'Installing dependecies. This might take a few minutes or more!';
	foreach my $d (@install) {
		my $cmd = qq~perl -MCPAN -e "force install $d" >> dep.log~;
		nfo $cmd;
		`$cmd`;
		nfo 'Done';
	}
	nfo 'Please, check dep.log for more info';
	return @install;
}

sub trim($) {
	my ($value) = @_;
	$value =~ s/^[\s\t]+//;
	$value =~ s/[\s\t]$//;
	return $value;
}

sub readExcell($) {
	my ($path) = @_;
	require Spreadsheet::ParseExcel;

	nfo "Reading $path";
	unless (-f $path) {
		nag "No such file $path";
		return;
	}

	my $parser   = Spreadsheet::ParseExcel->new();
	my $workbook = $parser->parse($path);

	unless (defined $workbook) {
		require Spreadsheet::XLSX;

		$workbook = Spreadsheet::XLSX->new($path);

		unless (defined $workbook) {
			nag $parser->error();
			return;
		}
	}

	my %varsInv;
	my @rows;
	my $totalRows = 0;
	for my $worksheet ($workbook->worksheets()) {
		my ($row_min, $row_max) = $worksheet->row_range();
		my ($col_min, $col_max) = $worksheet->col_range();

		for my $col ($col_min .. $col_max) {
			my $cell = $worksheet->get_cell($row_min, $col);
			my $value = trim($cell->unformatted());
			$varsInv{$col} = $value;
		}

		for my $row ($row_min + 1 .. $row_max) {
			my %rowData;
			for my $col ($col_min .. $col_max) {
				my $cell = $worksheet->get_cell($row, $col);
				next unless $cell;

				my $value = trim($cell->unformatted());
				$rowData{ $varsInv{$col} } = $value;
			}
			$totalRows++;
			push @rows, \%rowData;
		}
	}

	nfo "Found $totalRows rows";

	return { rows => \@rows };
}

sub readAllExcells($$) {
	my ($path, $all) = @_;
	my %result;
	foreach my $alias (keys %$all) {
		my $file = $all->{$alias};
		$result{$alias} = readExcell "$path/$file";
	}
	return \%result;
}

sub parseCondition {
	my ($str) = @_;
	$str =~ s/\s*==\s*/ eq /g;
	$str =~ s/\s*!=\s*/ ne /g;
	return "if ($str) {";
}

sub readTemplate($) {
	my ($path) = @_;
	my $txt = Eldhelm::Util::FileSystem->getFileContents($path);

	$txt =~ s/^[\t\s]*([^\s\t#].*)/"_writeToFile(qq~$1~);"/gme;
	$txt =~ s/^[\t\s]*#append to (.*)/"_setOutputFile(qq~$1~);"/gme;
	$txt =~ s/^[\t\s]*#if[\s\t]+(.*)/parseCondition($1)/igme;
	$txt =~ s/^[\t\s]*#endif/\}/igm;
	$txt =~ s/^[\t\s]*(#.*)/"_writeToFile(qq~$1~);"/gme;
	$txt =~ s/\$(\w+)/"\$_parsedData->{'$1'}"/ge;
	$txt =~ s/\$\{([^\{\}]+)\}/"\$_parsedData->{'$1'}"/ge;
	$txt .= "\n_setOutputFile();";
	$txt .= "\n1;";

	return $txt;
}

sub readAllTemplates($) {
	my ($all) = @_;
	my @result;
	foreach my $tpl (@$all) {
		push @result, readTemplate $tpl;
	}
	return \@result;
}

sub readConfig() {
	return do 'config.pl' or nag "Config error: $@";
}

my $_currentFH;
my $_outputFolder;
my %_usedFiles;

sub _setOutputFile {
	my ($file) = @_;
	if ($_currentFH) {
		nfo "OK";
		_writeToFile('#===================================================================')
			if $args{s} || $args{separator};
		close $_currentFH;
	}
	return unless $file;

	my $path = "$_outputFolder/$file";
	nfo(($_usedFiles{$path} ? 'Using' : '* Creating')." file $path");
	open $_currentFH, $_usedFiles{$path} || '>', $path;
	$_usedFiles{$path} = '>>';
}

sub _writeToFile {
	my ($content) = @_;
	unless ($_currentFH) {
		nag 'Can not write content';
		exit;
	}
	print $_currentFH "$content\n";
}

sub processTeplate {
	my ($tpl, $_parsedData) = @_;
	eval $tpl or do {
		my $i = 0;
		$tpl =~ s/^/$i++; " $i. "/gme;
		nag $tpl;
		nag $@;
		exit;
	};
}

my $cmd = Eldhelm::Util::CommandLine->new(
	argv    => \@ARGV,
	items   => ['folder to process'],
	options => [
		[ 'h help',      'help' ],
		[ 'c check',     'check dependencies' ],
		[ 'p process',   'folder to process' ],
		[ 'o output',    'output folder; defaults to output' ],
		[ 'd debug',     'prints compiled template' ],
		[ 's separator', 'appends a separator before file close' ]
	],
	examples => [ "perl $0 -p xls", "perl $0 xls" ]
);

%args = $cmd->arguments;

if ($args{h} || $args{help} || !keys %args) {
	print $cmd->usage;
	exit;
}

if ($args{check}) {
	checkDeps;
	exit;
}

my $folder = $args{p} || $args{process} || $args{list}[0];
unless (-d $folder) {
	nag "There is no such folder $folder";
	exit;
}

$_outputFolder = $args{o} || $args{output} || 'output';
unless (-d $_outputFolder) {
	nag "There is no such folder $_outputFolder";
	exit;
}

nfo 'Reading files ...';
my $config    = readConfig;
my $templates = readAllTemplates $config->{template};
if ($args{d} || $args{debug}) {
	nfo $_ foreach @$templates;
	exit;
}

my $excell = readAllExcells $folder, $config->{excell};

nfo 'Processing ...';
my $tpl = $templates->[0];

my %newServices;
foreach my $row (@{ $excell->{main}{rows} }) {
	my $dn = $row->{'OLD device name'};
	next unless $dn;
	my $k = $dn.'-'.$row->{'OLD device port'};
	$newServices{$k} = $row;
}

my %vlans;
foreach my $row (@{ $excell->{vlan}{rows} }) {
	$vlans{ $row->{'VLAN ID'} } = $row;
}

my %ips;
foreach my $row (@{ $excell->{ips}{rows} }) {
	$ips{ $row->{NE_NAME} } = $row;
}

my %rsvp;
foreach my $row (@{ $excell->{rsvp}{rows} }) {
	$rsvp{ $row->{ACR} } = $row;
}

my @errors;
my %uniqueServices;
foreach my $row (@{ $excell->{old}{rows} }) {
	next unless keys %$row;
	my $k      = $row->{devicename}.'-'.$row->{portname};
	my $newRow = $newServices{$k};
	my $nk     = $newRow->{'NEW device name'}.'-'.$row->{VLAN};

	$k = $row->{VLAN};
	unless ($k) {
		push @errors, "Vlan is blank for $row->{devicename} $row->{portname}!";
		next;
	}
	my $vlanRows = $vlans{$k};
	unless ($vlanRows) {
		push @errors, "Vlan $k not found for $row->{devicename} $row->{portname}!";
		next;
	}

	$k = $newRow->{'NEW device name'};
	my $ipRows = $ips{$k};
	unless ($ipRows) {
		push @errors, "New device $k not found in IP plan for $row->{devicename} $row->{portname}!";
		next;
	}

	my $rsvpRows = $rsvp{$k};
	unless ($rsvpRows) {
		push @errors, "New device $k not found in RSVP for $row->{devicename} $row->{portname}!";
		next;
	}

	$uniqueServices{$nk} = { %$row, %$newRow, %$vlanRows, %$ipRows, %$rsvpRows };
}

foreach (sort { $a cmp $b } keys %uniqueServices) {
	my $row = $uniqueServices{$_};
	if ($row->{'QoS marking'} eq 'not migrate') {
		nfo "QoS marking = not migrate, skipping $row->{devicename} $row->{portname}";
		next;
	}
	processTeplate($tpl, $row);
}

nag $_ foreach @errors;

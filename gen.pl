use strict;
use lib 'lib/';
use Carp;
use Data::Dumper;
use Eldhelm::Util::CommandLine;
use Eldhelm::Util::FileSystem;

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

	my %vars      = {};
	my %varsInv   = {};
	my %lists     = {};
	my %values    = {};
	my $totalRows = 0;
	for my $worksheet ($workbook->worksheets()) {
		my ($row_min, $row_max) = $worksheet->row_range();
		my ($col_min, $col_max) = $worksheet->col_range();

		for my $col ($col_min .. $col_max) {
			my $cell = $worksheet->get_cell($row_min, $col);
			my $value = trim($cell->unformatted());
			$varsInv{$col} = $value;
			$vars{$value}  = $col;
		}

		for my $row ($row_min + 1 .. $row_max) {
			for my $col ($col_min .. $col_max) {
				my $cell = $worksheet->get_cell($row, $col);
				next unless $cell;

				my $value = trim($cell->unformatted());

				$lists{ $varsInv{$col} } ||= [];
				$lists{ $varsInv{$col} }[ $row - 1 ] = $value;

				$values{ $varsInv{$col} } ||= {};
				$values{ $varsInv{$col} }{$value} = $row;
			}
			$totalRows++;
		}
	}

	nfo "Found $totalRows rows";

	return {
		vars    => \%vars,
		varsInv => \%varsInv,
		lists   => \%lists,
		values  => \%values,
	};
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

sub readTemplate($) {
	my ($path) = @_;
	my $txt = Eldhelm::Util::FileSystem->getFileContents($path);

	$txt =~ s/^[\t\s]*([^\s\t#].*)/"_writeToFile(qq~$1~);"/gme;
	$txt =~ s/^[\t\s]*#append to (.*)/"_setOutputFile(qq~$1~);"/gme;
	$txt =~ s/^[\t\s]*#if[\s\t]+(.*)/"if($1){"/igme;
	$txt =~ s/^[\t\s]*#endif/\}/igm;
	$txt =~ s/\$\w/"\$_parsedData->{'$1'}"/ge;
	$txt =~ s/\$\{([^\{\}]+)\}/"\$_parsedData->{'$1'}"/ge;
	$txt .= "_setOutputFile();";
	$txt .= "1;";

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
	return do 'config.pl';
}

my $_currentFH;
my $_outputFolder;
my %_usedFiles;

sub _setOutputFile {
	my ($file) = @_;
	if ($_currentFH) {
		nfo "OK";
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
	print $_currentFH "$content\n";
}

sub processTeplate {
	my ($tpl, $_parsedData) = @_;
	eval $tpl or do {
		my $i = 0;
		$tpl =~ s/^/$i++; " $i. "/gme;
		nag $tpl;
		nag $@;
	};
}

my $cmd = Eldhelm::Util::CommandLine->new(
	argv    => \@ARGV,
	items   => ['folder to process'],
	options => [
		[ 'h help',    'help' ],
		[ 'c check',   'check dependencies' ],
		[ 'p process', 'folder to process' ],
		[ 'o output',  'output folder; defaults to output' ]
	],
	examples => [ "perl $0 -p xls", "perl $0 xls" ]
);

my %args = $cmd->arguments;

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
my $excell    = readAllExcells $folder, $config->{excell};

nfo 'Processing ...';
my $tpl = $templates->[0];

foreach my $port (@{ $excell->{old}{lists}{'Port description'} }) {
	next unless $port;
	next unless $excell->{main}{values}{'Port Description'}{$port};

	nag $port;

	processTeplate($tpl, {
		'NEW device name' => $port
	});
}


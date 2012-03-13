#!/usr/bin/perl

use strict;
use Spreadsheet::ParseExcel;

my $file_name = 'test.xls';
my $spreadsheet_number = 0;
my $delimiter = '|';
my $line_ending = "\n";

my $parser = Spreadsheet::ParseExcel->new();
my $workbook = $parser->parse($file_name);

if ( !defined $workbook ) { die $parser->error(), ".\n" }

my $worksheet = $workbook->worksheet($spreadsheet_number);

my ($row_min, $row_max) = $worksheet->row_range();
my ($col_min, $col_max) = $worksheet->col_range();

for my $row ($row_min..$row_max) {
	for my $col ($col_min..$col_max) {
		my $cell = $worksheet->get_cell($row,$col);
		if (!$cell) {
			print "";
		} else {
			print $cell->value();
		}
		print $delimiter unless ($col == $col_max);
	}
	print $line_ending;
}


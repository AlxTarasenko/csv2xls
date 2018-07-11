#!/usr/bin/perl -w
#
# Version 1.5
#
# Usage: csv2xls.pl file.csv newfile.xls
#
# The program uses Text::CSV_XS to parse the CSV.
# First string, this title of columns, for set filter on column
# Title suffix: "=" align left, "+" align center, "-" align right
# Max width Column is bigger of: $cell_max_width OR $word_max_width
# Shrink width Column by max width of phrase placed in Max width Column
#
# 06/2018, Alexey Tarasenko, atarasenko@mail.ru
#
# Red Hat:
# yum install perl-Spreadsheet*
# yum install perl-Text-CSV_XS*
#

use strict;
use Spreadsheet::WriteExcel;
use Text::CSV_XS;

# Check for valid number of arguments
if ( ( $#ARGV < 1 ) || ( $#ARGV > 2 ) ) {
    die("Usage: csv2xls csvfile xlsfile\n");
    }
    
# Open the Comma Separated Variable file
open( CSVFILE, $ARGV[0] ) or die "$ARGV[0]: $!";
    
# Create a new Excel workbook
my $workbook  = Spreadsheet::WriteExcel->new( $ARGV[1] );
my $worksheet = $workbook->add_worksheet();

my $format_cell_left = $workbook->add_format();
$format_cell_left->set_font('Courier New');
$format_cell_left->set_num_format('@');
$format_cell_left->set_valign('top');
$format_cell_left->set_text_wrap();
$format_cell_left->set_align('left');

my $format_cell_center = $workbook->add_format();
$format_cell_center->set_font('Courier New');
$format_cell_center->set_num_format('@');
$format_cell_center->set_valign('top');
$format_cell_center->set_text_wrap();
$format_cell_center->set_align('center');

my $format_cell_right = $workbook->add_format();
$format_cell_right->set_font('Courier New');
$format_cell_right->set_num_format('@');
$format_cell_right->set_valign('top');
$format_cell_right->set_text_wrap();
$format_cell_right->set_align('right');

my $format_cell = $workbook->add_format();
$format_cell->set_font('Courier New');
$format_cell->set_num_format('@');
$format_cell->set_valign('top');
$format_cell->set_text_wrap();

my $format_hdr = $workbook->add_format();
#$format_hdr->set_font('Courier New');
#$format_hdr->set_num_format('@');
$format_hdr->set_bold();
$format_hdr->set_bg_color('yellow');
$format_hdr->set_align('center');
$format_hdr->set_align('vcenter');
#$format_hdr->set_color('red');
#$format_hdr->set_bg_color('grey');
#$format_hdr->set_size(11);

# Rows and Columns counters
my $cols = 0;
my $rows = 0;
# Array of width columns
my @col_width = ();
$#col_width = -1;
# Array of align columns
my @col_align = ();
$#col_align = -1;
# Max cell width
my $cell_max_width = 50;
# The width corresponds to the column width value that is specified in Excel. 
# It is approximately equal to the length of a string in the default font of Arial 10. 
# Unfortunately, there is no way to specify "AutoFit" for a column in the Excel file format. 
# This feature is only available at runtime from within Excel.
my $cel_mul = 1.15;
    
# Create a new CSV parsing object
my $csv = Text::CSV_XS->new;

while (<CSVFILE>) {
    if ( $csv->parse($_) ) {
        my @Fld = $csv->fields;
            
        my $col = 0;
        foreach my $token (@Fld) {
    	    if ( $rows == 0 ) {
		$col_align[ $col ] = "";
		if ( length($token) >= 1 ) {
		    my $char = substr( $token, -1, 1 );
		    if ( $char eq "=" ) { $col_align[ $col ] = "left"; $token = substr( $token, 0, (length($token)-1) ); }
		    if ( $char eq "+" ) { $col_align[ $col ] = "center"; $token = substr( $token, 0, (length($token)-1) ); }
		    if ( $char eq "-" ) { $col_align[ $col ] = "right"; $token = substr( $token, 0, (length($token)-1) ); }
		}
    	    }
	    my $len = length( $token ) + 2; # add 1 char at left and right sidies
	    if ( $len > $cell_max_width ) { 
		# if string length more MAX, limit by MAX
		$len = $cell_max_width;

		# calc max Phrase (0..MAX) and one Word length
		my $word_max_width  = 0;
		my $phrase_max_width = 0;
		my @words=split(' ',$token);
		my $phrase_width = 0;
		foreach my $word (@words) {
		    # one word and finish spase
		    my $word_length = length($word) + 1;
		    
		    # calc phrase lentgh
		    if ( ($phrase_width + $word_length) > $cell_max_width ) {
			# calc max phrase
			if ( $phrase_width > $phrase_max_width ) { $phrase_max_width = $phrase_width; }
			# set phrase current word
			$phrase_width = $word_length;
		    } else {
			$phrase_width = $phrase_width + $word_length;
		    }
		    # if one word, correct $phrase_max_width
		    if ( $phrase_max_width == 0 ) { $phrase_max_width = $word_length; }
		    
		    # calc max word
		    if ( $word_max_width < $word_length ) { $word_max_width = $word_length; }
		}
		if ( $phrase_max_width < $len ) { $len = $phrase_max_width; }
		if ( $word_max_width > $len ) { $len = $word_max_width; }
	    }
	    if ( exists $col_width[ $col ] ) {
		if ( $len > $col_width[ $col ] ) { $col_width[ $col ] = $len; }
	    } else {
		$col_width[ $col ] = $len;
	    }
            $col++;
	}
        $rows++;
    } else {
        my $err = $csv->error_input;
        print "Text::CSV_XS parse() failed on argument: ", $err, "\n";
    }
}
# Set the file pointer to start position
seek CSVFILE, 0, 0;

# Set filter for all columns
$cols = $#col_width;
$rows--;
$worksheet->autofilter(0, 0, $rows, $cols);
# Set format for all cells
$worksheet->set_column(0, $cols, undef, $format_cell);
# Set width for columns
for( my $i = 0; $i <= $#col_width; $i++ ) { 
    my $width = int( $cel_mul * $col_width[ $i ] );
    $worksheet->set_column( $i, $i, $width ); 
}
# Set align for columns
for( my $i = 0; $i <= $#col_align; $i++ ) { 
    my $align = $col_align[ $i ];
    if ( $align eq "left"   ) { $worksheet->set_column( $i, $i, undef, $format_cell_left ); }
    if ( $align eq "center" ) { $worksheet->set_column( $i, $i, undef, $format_cell_center ); }
    if ( $align eq "right"  ) { $worksheet->set_column( $i, $i, undef, $format_cell_right ); }
}

# Row and column are zero indexed
my $row = 0;
    
while (<CSVFILE>) {
    if ( $csv->parse($_) ) {
        my @Fld = $csv->fields;
            
        my $col = 0;
        foreach my $token (@Fld) {
	    #$worksheet->write( $row, $col, $token );
            if ( $row == 0 ) {
		if ( length($token) >= 1 ) {
		    my $char = substr( $token, -1, 1 );
		    if ( $char eq "=" || $char eq "+" || $char eq "-" ) { $token = substr( $token, 0, (length($token)-1) ); }
		}
        	$worksheet->write_string( $row, $col, $token, $format_hdr );
            } else { 
        	#$worksheet->write_string( $row, $col, $token, $format_str );
        	$worksheet->write_string( $row, $col, $token );
            }
            $col++;
	}
        $row++;
    } else {
        my $err = $csv->error_input;
        print "Text::CSV_XS parse() failed on argument: ", $err, "\n";
    }
}

#  Perl scripts For Language International 
#  Author: Andrew.yin 2016/5/15

#!/usr/bin/perl
use strict;
use Encode;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtUnicode;

# "CP936" can read chinese but Arabic is failed
# //TODO
my $oFmtC = Spreadsheet::ParseExcel::FmtUnicode->new(Unicode_Map=>"CP936"); 
# my $oFmtC = Spreadsheet::ParseExcel::FmtUnicode->new(); 
my $parser = Spreadsheet::ParseExcel->new();
my $workbook = $parser->parse('strings.xls', $oFmtC);
if ( !defined $workbook ) {
    die $parser->error(), ".\n";
}

my $dirPrefixName = "values-";
my $fileName = "/string.xml";

my $statementPrefix = "<string name=\"";
my $statementPrefixEnd = "\">";
my $statementSuffix = "</string>";

# If string.xml has change the style, you can change here to adapter
my $xmlResourceStart = "<resources>\n";
my $xmlResourceEnd = "</resources>";

for my $worksheet ( $workbook->worksheets() )
{
    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();

    our @dirName;
    our @stringNameArray;
    our $num = 0;
    for my $row ( $row_min .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {
            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;
            # if statement not null
            if($cell->value()) {

            	# if row = 0 , if dir not make, make string dir
            	if($row eq "0" && $col ne "0") {
            		$dirName[$col - 1] = $dirPrefixName.$cell->value();
					mkdir($dirName[$col - 1]) || die "Mkdir failed, Please delete Folder then exectue", "\n";
            		# open file and add title
            		&createHead($dirName[$col - 1]);
            	}

            	# add string name to array 	
        		if($row ne "0" && $col eq "0") {
        			$stringNameArray[$num] = $cell->value();
        			$num = $num + 1;
        		}

				# open file add string 
        		if($row ne "0" && $col ne "0") {
					my $filePath = $dirName[$col - 1].$fileName;
    				open(my $fhd, ">>", $filePath) or die "Fail to open the file $! \n";
    				my $str = &stringValue($stringNameArray[$row - 1], $cell->value());
    				say $fhd encode_utf8(decode("gbk", $str));
    				close $fhd
        		}
        	}
        	
        }
    }

    # add </resources>
    for my $i (@dirName) {
		my $filePath = $i.$fileName;
		open(my $fhd, ">>", $filePath) or die "Fail to open the file $! \n";
		say $fhd encode_utf8($xmlResourceEnd);
		close $fhd
	}

	print "\n ============== Successful ! =========== \n";
}

sub stringValue {
	"\t".$statementPrefix.$_[0].$statementPrefixEnd.$_[1].$statementSuffix."\n";
}

# Add Head <resources>
sub createHead {
	my $filePath = $_[0].$fileName;
	open(my $fhd, ">>", $filePath) or die "Fail to open the file $! \n";
	say $fhd encode_utf8($xmlResourceStart);
	close $fhd
}
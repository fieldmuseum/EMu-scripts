#!/usr/bin/perl -w
#
# DESCRIPTION: The convertSchemaToExcel.pl script is used to convert the schema.pl file to an Excel file format
#

#
# USE VARIABLES
#
use 	utf8;
use 	strict;
use 	warnings;
use 	Data::Dumper;
use 	IO::Handle;
use 	Spreadsheet::WriteExcel;
use 	List::UtilsBy qw(max_by);

#
# REQUIRE VARIABLES
#
require	"$ENV{EMUPATH}/utils/schema.pl";

#
# GLOBAL VARIABLES
#
our 	%Schema;
my	%headerColumns = ();
my 	$outputFile;
my 	$workbook;
my 	$worksheet;
my 	@headerValues = ();
my 	%headerStyle;
my 	%headerShading;
my 	$greyRGB;
my 	$headerFrmt;
my 	%style;
my 	$lightBlueRGB;
my 	%shading;
my 	$normalFrmt;
my 	$shadedFrmt;
my 	$currentFrmt;
my 	$addShading;
my 	$row;

#
# START SCRIPT
#

#
# Check user input count
#
if (@ARGV != 1)
{
        die("$0 [excel_file_name]\n");
}

#
# Assign outputFile
#
$outputFile = shift;
if ($outputFile !~ /\.xls$/)
{
        $outputFile .= ".xls";
}

#
# FUNCTION CALLS
#
print("Started schema.pl to Excel conversion...\n");
ReadSchema();
NewExcelDocument();
GenerateSchemaList();
CleanUp();

print("...DONE\n");

#
# EXIT SCRIPT
#
exit(0);


#
# FUNCTION DECLARATIONS
#

#
# The ReadSchema function is used to get the header column names
#
sub
ReadSchema
{
	my $columns;
	my $properties;
	my $value;
	my $index = 0;

	$headerColumns{Table} = $index;
	$headerValues[$index] = "Table";
	foreach my $table (sort (keys %Schema))
	{
		$columns = $Schema{$table}->{columns};
		foreach my $column (sort (keys %{$columns}))
		{
			$properties = $columns->{$column};
			foreach my $property (sort (keys %{$properties}))
			{
				$value = $properties->{$property};
				if (! defined($headerColumns{$property}))
				{
					$index++;
					$headerColumns{$property} = $index;
					$headerValues[$index] = $property;
				}
			}
		}
	}
}

#
# The NewExcelDocument function is used to create a blank Excel file
#
sub
NewExcelDocument
{
	my $headerCol = 0;
	
	# 
	# Spreadsheet position variables
	#
	$row = 0;
	
	#
	# Create workbook
	#
	$workbook = Spreadsheet::WriteExcel->new($outputFile);
	$worksheet = $workbook->add_worksheet('Mapping to EMu');
	$worksheet->set_landscape();
	$worksheet->set_margins(0.2);
	$worksheet->set_header('', 0.1);
	$worksheet->set_footer('', 0.1);
	$worksheet->center_horizontally();
	$worksheet->center_vertically();

	#
	# Header formatting
	#
	%headerStyle = (
			font    => 'Arial',
			size    => 10,
			bold    => 1,
		);
	$greyRGB = $workbook->set_custom_color(40, 216, 218, 219);
	%headerShading = (
			bg_color => $greyRGB,
			pattern  => 1,
		);
	$headerFrmt = $workbook->add_format(%headerStyle, %headerShading);

	# 
	# Print any header information here on the first row of the file
	#
	foreach my $headerValue (@headerValues)
	{
	        $worksheet->write($row, $headerCol, $headerValue, $headerFrmt);
	        $headerCol++;
	}
	$row++;

	#
	# Workbook body formatting
	#
	%style = (
			font    => 'Arial',
	                size    => 10,
	        );
	$lightBlueRGB = $workbook->set_custom_color(41, 180, 198, 231);
	%shading = (
                    	bg_color => $lightBlueRGB,
                        pattern  => 1,
                  );
	$normalFrmt = $workbook->add_format(%style);
	$shadedFrmt = $workbook->add_format(%style, %shading);		
	$addShading = 0;
}

#
# The GenerateSchemaList if a function which will add the schema data to the Excel workbook
#
sub
GenerateSchemaList
{
	my $tableData;
	my $properties;
	my $columns;
	my $value;
	my $isModified;
	my $rowMax;
	my $fieldLength;
	my $nestLength;

	foreach my $table (sort (keys %Schema))
	{
		$isModified = 0;
		if($addShading)
	        {
	                $currentFrmt = $shadedFrmt;
	        }
	        else
	        {
	        	$currentFrmt = $normalFrmt;
	        }

		$columns = $Schema{$table}->{columns};
		foreach my $column (sort (keys %{$columns}))
		{
			$properties = $columns->{$column};
			foreach my $property (keys %headerColumns)
			{
				if ($property eq "Table")
				{
					$worksheet->write($row, $headerColumns{'Table'}, $table, $currentFrmt);
				}
				elsif (defined($properties->{$property}))
				{
					$value = $properties->{$property};
					if ($property eq "ItemFields")
					{
						$fieldLength = "";
						foreach my $tableRow (@{$value})
						{
							eval
							{
								$nestLength = "";
								foreach my $nestRow (@{$tableRow})
								{
									if ($nestLength ne '')
									{
										$nestLength .= ",";
									}
									$nestLength .= $nestRow;
								}
								if ($fieldLength ne '')
								{
									$fieldLength .= " | ";
								}
								$fieldLength .= $nestLength;
							};
							if ($@)
							{
								if ($fieldLength ne '')
								{
									$fieldLength .= ", ";
								}
								$fieldLength .= $tableRow;
							}
						}

						$value = $fieldLength;
					}
					$worksheet->write($row, $headerColumns{$property}, $value, $currentFrmt);
				}
				else
				{
					$worksheet->write($row, $headerColumns{$property}, "", $currentFrmt);
				}
			}
			
			$row++;
			$isModified = 1;
		}

		if ($isModified)
		{
			$addShading = !$addShading;
		}
	}
}

#
# The CleanUp function cleans up any outstanding processes
#
sub
CleanUp
{
	if ($workbook)
	{
		$workbook->close() or die "Error closing file: $!";
	}
}

#
# END OR SCRIPT - SHOULD NEVER REACH THIS POINT
#
exit(1);

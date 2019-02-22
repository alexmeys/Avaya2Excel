#!/usr/bin/perl
# av2excel.pl
#Module version Activestate Perl: 
#Mime-lite (3.028)
#Excel-writer-xlsx (0.45)
#Net::SMTP 
#Alex Meys

use warnings;
#use strict;
use File::Copy;
use Excel::Writer::XLSX;
use MIME::Lite;
use Net::SMTP;

#Create variables
my $file = "C:\\avaya\\raw\\raw.log";
my $line = "";
my @arr = ();
my @words = ();
my $i=1;
my $j=0;
(my $second, my $minute, my $hour, my $mday, my $mon, my $year)  =localtime(); 
$year = $year+1900;
$mon = $mon+1;
my $timestamp = "_".$year."_".$mon."_".$mday;

#Create an Excel file
my $workbook = Excel::Writer::XLSX->new("C:\\avaya\\excel\\avaya$timestamp.xlsx");
my $sheet = $workbook->add_worksheet("Telephony");

#Make it look nice in Excel
my $opmaak = $workbook->add_format();
$opmaak->set_color('blue');
$opmaak->set_align('center');
my $opmaak2 = $workbook->add_format();
$opmaak2->set_align('left');
my $opmaak3 = $workbook->add_format();
$opmaak3->set_bold();
my $opmaaktbl = $workbook->add_format();
$opmaaktbl->set_border(1);
$opmaaktbl->set_align('left');

#Create manual headers for columns
$sheet->write(0, 0, "Start time", $opmaak);
$sheet->write(0, 1, "Duration", $opmaak);
$sheet->write(0, 2, "Rings", $opmaak);
$sheet->write(0, 3, "caller", $opmaak);
$sheet->write(0, 4, "In or outbound", $opmaak);
$sheet->write(0, 5, "Number called", $opmaak);
$sheet->write(0, 6, "Username / linenr", $opmaak);
$sheet->write(0, 7, "Hold Time", $opmaak);
$sheet->write(0, 8, "Park Time", $opmaak);

#open raw logfile
open(LOG, "$file") || die "Cannot open file\n";

#The actual work, loop starts here, from log to Excel table.
my @lines = <LOG>;
foreach $line (@lines)
{
    @words = split(',',$line);
	$words[0] =~ s/^\s+//;
	$words[0] =~ s/\s+$//;
	$sheet->write_string($i, $j, "$words[0]", $opmaak2);
	$j +=1;
	$words[1] =~ s/^\s+//;
	$words[1] =~ s/\s+$//;
	$sheet->write_string($i, $j, "$words[1]", $opmaak2);
	$j +=1;
	$words[2] =~ s/^\s+//;
	$words[2] =~ s/\s+$//;
	$sheet->write($i, $j, "$words[2]", $opmaak2);
	$j +=1;
	$words[3] =~ s/^\s+//;
	$words[3] =~ s/\s+$//;
	$sheet->write_string($i, $j, "$words[3]", $opmaak2);
	$j +=1;
	$words[4] =~ s/^\s+//;
	$words[4] =~ s/\s+$//;
	$sheet->write_string($i, $j, "$words[4]", $opmaak2);
	$j +=1;
	$words[5] =~ s/^\s+//;
	$words[5] =~ s/\s+$//;
	$sheet->write_string($i, $j, "$words[5]", $opmaak2);
	$j +=1;
	$words[6] =~ s/^\s+//;
	$words[6] =~ s/\s+$//;
	$sheet->write($i, $j, "$words[6]", $opmaak2);
	$j +=1;
	$words[7] =~ s/^\s+//;
	$words[7] =~ s/\s+$//;
	$sheet->write($i, $j, "$words[7]", $opmaak2);
	$j +=1;
	$words[8] =~ s/^\s+//;
	$words[8] =~ s/\s+$//;
	$sheet->write($i, $j, "$words[8]", $opmaak2);
	$j =0;
	$i +=1;	
}

#Add your own users
my @gebruikers = ("user1", "user2", "user3", "user4", "user5", "user6", "user7", "user8","usera", "userb", "userc", "userd", "usere", "userf", "userg", "userh","useri", "userj", "userk", "userl", "userm", "usern", "usero", "userp");

my $usernms = @gebruikers;
$usernms = ($usernms-1);


#Add letters for excel
my @letters = ('m'..'xfd');
my $it = 0;

#Write down the users and count them for cell usage
foreach my $naam (@gebruikers)
{
  $sheet->write($letters[$it].'1', $gebruikers[$it], $opmaak3);
  $it++;
}

#Default titles fixed position
$sheet->write('J4', "In", $opmaak3);
$sheet->write('J6', "Out", $opmaak3);
$sheet->write('K1', "Totals", $opmaak3);
$sheet->write('J8', "Max", $opmaak3);
$sheet->write('J9', "Min", $opmaak3);
$sheet->write('J10', "Avg", $opmaak3);

#Find how many times a users occures in a row
my $it2 = 0;
#Enter here total of phonecalls calculations will be done on. 
#(if you run this program daily/weekly and < 5000 phonecalls will be placed during this week or day period, your good.)

my $trecords = 5000;

foreach my $nam (@gebruikers)
{
  $sheet->write_formula($letters[$it2].'2', "=COUNTIF(G2:G".$trecords.",$letters[$it2]1)", $opmaak2);
  $it2++;
}


#Incomming calls count
my $it3 = 0;
foreach my $nm (@gebruikers)
{
  $sheet->write_formula($letters[$it3].'4', "=SUMPRODUCT((E2:E".$trecords."=\"I\")*(G2:G".$trecords."=\"$gebruikers[$it3]\"))",$opmaak2);
  $it3++;
}

#Outgoing calls count
my $it4 = 0;
foreach my $nom (@gebruikers)
{
  $sheet->write_formula($letters[$it4].'6', "=SUMPRODUCT((E2:E".$trecords."=\"O\")*(G2:G".$trecords."=\"$gebruikers[$it4]\"))", $opmaak2);
  $it4++;
}

#Totals and basic caluclations for displaying. Fixed position
$it4 = $it4-1;
$sheet->write('K2', "=SUM(M2:$letters[$it4]2)", $opmaaktbl);
$sheet->write('K4', "=SUM(M4:$letters[$it4]4)", $opmaaktbl);
$sheet->write('K6', "=SUM(M6:$letters[$it4]6)", $opmaaktbl);
$sheet->write('K8', "=MAX(M2:$letters[$it4]2)", $opmaaktbl);
$sheet->write('K9', "=MIN(M2:$letters[$it4]2)",$opmaaktbl);
$sheet->write('K10', "=AVERAGE(M2:$letters[$it4]2)",$opmaaktbl);

my $end = $letters[$usernms];
$end = uc($end);

my $in1_max = "$end"."1";
my $in2_max = "$end"."4";
my $uit1_max = "$end"."1";
my $uit2_max = "$end"."6";

#Create graph not embedded, on new page with name Graph
my $chart = $workbook->add_chart(type => 'column', name => 'Graph');

$chart->add_series(
	name => 'In',
	categories => "=Telephony!M1:$in1_max",
	values	=> "=Telephony!M4:$in2_max",
);
$chart->add_series(
	name => 'Out',
	categories => "=Telephony!M1:$uit1_max",
	values => "=Telephony!M6:$uit2_max",
);

#Close log and move it for archiving.
$workbook->close();
close(LOG);
my $newloc = "C:\\avaya\\old\\raw$timestamp.log";
move($file,$newloc);

#start mail function (need MIME::Lite and Net::SMTP sending)

my $report_file = "C:\\avaya\\excel\\avaya$timestamp.xlsx";

my $msg = MIME::Lite->new(
    #From adres, eg: phonemonitor@company.area (this can be fake, but do check spam...)
    From    => 'Phonemonitor@mycompany.area',
	#To who should we send this mail?
    To      => 'myself@company.area',
    Subject => "Excel Report Telephony usage",
    Type    => 'multipart/mixed'
);
$msg->attach(
    Type => 'text/plain',
	#Change text if wanted, change YOUR NAME and COMPANY with your credentials. This will be in the message body
    Data => "Dear,\n\nIn attachment you will find the telephony statistics.\n\nKind regards,\nYOUR NAME\nCOMPANY.",
);
$msg->attach(
    Type        => 'application/zip',
    Path        => $report_file,
    Disposition => 'attachment',
);
#set the outgoing server of your ISP
$msg->send('smtp','your.outgoing.server');

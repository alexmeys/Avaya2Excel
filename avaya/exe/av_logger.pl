#!/usr/bin/perl
# Alex Meys
# av_logger.pl
#/*==================================


use IO::Socket::INET;
use strict;
use warnings;

sub verbinding
{
  my $sock = new IO::Socket::INET (
    #Edit peerhost with the PBX IP
    PeerHost => '192.168.1.200',
    #Port can be changed if not free on the system
    PeerPort => '8087',
    Proto => 'tcp',
	Reuse => 1) || die "\[-\]Problem with socket connection : $!\n";
  return $sock;
}

sub looping
{
  my $sock = &verbinding();
  my @tel = ();
  my @words = ();
  my $output = "c:\\avaya\\raw\\raw.log";

  print "\n\[*\] Connected, Do not interrupt...\n";
  sleep 2;
  print "\[*\] ...recording...\n";

  while(<$sock>)
  {
    push(@tel, $_);
    if($#tel+1 < 1)
    {
      next;
    }
    else
    {
      open(F1, ">>$output");
	  foreach my $text (@tel)
	  {
	    @words = split(',',$text);
	    #If word matches autoattend (we use this for voicemail pro, so no inclusions of voicemails), skip it.
	    #If you have a voicemail pro, change the name to the right attendant name. Or remove in case you want them included.
	    if($words[5] =~ /autoattend/i)
	    {
	      next;
	    }
	    if($words[3] eq "")
	    {
	      $words[3] = "Suppressed";
	    }
	    if($words[1] eq "00:00:00")
	    {
	      next;
	    }
	    print F1 "$words[0] , $words[1] , $words[2] , $words[3] , $words[4] , $words[5] , $words[12] , $words[15], $words[16]\n";
	  }
	  close(F1);
	  @tel = ();
	  @words = ();
    }
  }$sock->close();
}

&verbinding();
&looping();

my $i = 1;

while($i)
{
  print "\n\n\[*\] Closing connection with PBX...\n";
  print "\[*\] Average time to reconnect 2 minutes.\n";
  sleep 150;
  print "\n\n";
  &verbinding();
  &looping();
}

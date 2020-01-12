##############################
# ENABLE MODULES
##############################
use Win32::OLE;
use Win32::OLE::Variant;
use Getopt::Std;



########################################################
# CHECK FOR SUFFICIENT NUMBER OF COMMAND LINE ARGUMENTS
########################################################
if ( !( (scalar(@ARGV) == 4) or (scalar(@ARGV) == 6) )) { die "

       +============================================+
       |   gnmapper.pl - Version 1.0	            |
       |       	By: Rajat Swarup	   	    |
       +============================================+

       ** Note: To be run solely from a MS Windows command prompt

Usages:	perl gnmapper.pl [-g nmap_portscan_result_file] [-o outputfile.xls]

	-g  [Grepable Nmap File]
	-o  [Output File]  (* must have a .xls extension)

";}




###############################
########## MAIN ###############
###############################

&process_cmd_line();	# Take in command arguments and initialize certain variables accordingly
&excel_setup();		# Necessary MS Excel initializations

if ($scan_type eq 'g') { &process_gnmap(); }		# Process Grepable Nmap File


&print_results();	# Insert results into excel spreadsheet

&finish_excel();	# Finish column/cell adjustments, save, and close out the new excel file

print "\n\t\*\* Job Complete \*\*\n";
exit;								


##############################################################################################
##################################### SUBROUTINES ############################################
##############################################################################################

sub process_cmd_line()
{
	# Grab Flags and their respective values from the command line.
	getopt('ngsro');
	
	# Ensure one and ONLY one input/port-scan-result flag has been enabled
	if ( !((defined ($opt_n)) xor (defined ($opt_g)) xor (defined ($opt_s))) ) {
		print "\n\t\tERROR: Incorrect Portscan Input Flag Usage\n";
		exit;
	}
	
	# The 'r' flag can only be use with the 's' flag
	# Ensure that the 'r' flag is used soley in conjunction with the 's' flag
	if ( defined $opt_r) {
		if ( !( defined $opt_s ) ) {
			# Flag 'r' is being used without flag 's'
			print "\n\t\tERROR: Where\'z da \'s\' flag fool\!\!\!\n";
			exit;
		}else { open(INPUT2, "<$opt_r") or die "\n\tERROR: Could not open superscan file: $opt_r. $!\n"; }
	}
		
	
	# Ensure that an output file has been specified and meets requirements.
	if (defined $opt_o) {
		$output_file = $opt_o;
		
		# Make sure the output file has an .xls extension
		if ( !($output_file =~ "\.xls") ) 
		{
			print "\n\tERROR: The ouput file\: $output_file should be named as a \"\.xls\" file\.  Thank you\.\n";
			exit;
		} 
		
	}else {
		# If this point is reached, $opt_o was not defined and ergo no output file was specified
		print "\n\t\tERROR: Please specify an output file\n";
		exit;
	}
	
	
	# At this point we know the user has specified an input and output file.
	# Now open the input file for reading:
	if (defined $opt_g) {
		$scan_type = 'g';
		open(INPUT, "<$opt_g") or die "\n\tERROR: Could not open gnmap file: $opt_g. $!\n";
	}

}


########################################################################	
sub excel_setup{

	# Necessary Excel setup code
	$excel = Win32::OLE->new('Excel.Application') or die "Cannot open Excel: $!\n"; # start Excel OLE  Server
	$new_workbook = $excel->Workbooks->add or die "Cannot open $current_dir: $!\n"; #open new excel workbook
	$excel->{DisplayAlerts} = 0;		# Turns off anoying alerts
	
	# Prepare worksheet to be used
	$sheet_one = $new_workbook->Worksheets("Sheet1");	
	$sheet_one->Activate(); #set to active worksheet
	$sheet_one->{name} = "Target Matrix";
	
	# Clean up extra sheets
	#$new_workbook->Worksheets("Sheet2")->Delete;	# Delete Sheet2
	#$new_workbook->Worksheets("Sheet3")->Delete;	# Delete Sheet3
	
	# Work will be saved in current directory 
	chomp($current_path = `cd`);			# grabs current working directory, chomps off its newline character
	$excel_doc = $current_path . "\\$output_file";	# Resulting excel document will be saved			
							# in the current directory under the name you specified
	
	# PRINT HEADER OF OUTPUT FILE
	$sheet_one->Range("a1")->{Value} = 'IP Address';	
	$sheet_one->Range("b1")->{Value} = "Fully\-qualified\nDomain Name";	
	$sheet_one->Range("c1")->{Value} = 'Operating System Guess';	
	$sheet_one->Range("d1")->{Value} = 'FTP (21)';
	$sheet_one->Range("e1")->{Value} = 'Telnet (23)';
	$sheet_one->Range("f1")->{Value} = 'SMTP (25)';
	$sheet_one->Range("g1")->{Value} = 'Domain (53)';
	$sheet_one->Range("h1")->{Value} = 'HTTP (80)';
	$sheet_one->Range("i1")->{Value} = 'Netbios (139)';
	$sheet_one->Range("j1")->{Value} = 'HTTPS (443)';
	$sheet_one->Range("k1")->{Value} = 'SSH (22)';
	$sheet_one->Range("l1")->{Value} = 'DHCPD (67)';
	$sheet_one->Range("m1")->{Value} = 'TFTP (69)';
	$sheet_one->Range("n1")->{Value} = 'Rpcbbind (111)';
	$sheet_one->Range("o1")->{Value} = 'Identd (113)';
	$sheet_one->Range("p1")->{Value} = 'NTP (123)';
	$sheet_one->Range("q1")->{Value} = 'MSRPC (135)';
	$sheet_one->Range("r1")->{Value} = 'SNMP (161)';
	$sheet_one->Range("s1")->{Value} = 'Http-Mgmt (280)';
	$sheet_one->Range("t1")->{Value} = 'LDAP (389)';
	$sheet_one->Range("u1")->{Value} = 'microsoft-ds (445)';
	$sheet_one->Range("v1")->{Value} = 'ISAKMP (500)';
	$sheet_one->Range("w1")->{Value} = 'Lotus Notes (1352)';
	$sheet_one->Range("x1")->{Value} = 'MS-SQL (1433)';
	$sheet_one->Range("y1")->{Value} = 'Citrix (1494)';
	$sheet_one->Range("z1")->{Value} = 'Oracle (1521)';
	$sheet_one->Range("aa1")->{Value} = 'Oracle (1527)';
	$sheet_one->Range("ab1")->{Value} = 'PPTP (1723)';
	$sheet_one->Range("ac1")->{Value} = '1776';
	$sheet_one->Range("ad1")->{Value} = '2701';
	$sheet_one->Range("ae1")->{Value} = 'symantec-av (2967)';
	$sheet_one->Range("af1")->{Value} = 'MySQL (3306)';
	$sheet_one->Range("ag1")->{Value} = 'RDP (3389)';
	$sheet_one->Range("ah1")->{Value} = 'Session Initiation (5060)';
	$sheet_one->Range("ai1")->{Value} = 'Oracle SQL Interface (5560)';
	$sheet_one->Range("aj1")->{Value} = 'PC Anywhere (5631)';
	$sheet_one->Range("ak1")->{Value} = 'VNC-HTTP (5800)';
	$sheet_one->Range("al1")->{Value} = 'VNC (5900)';
	$sheet_one->Range("am1")->{Value} = 'X11 (6000)';
	$sheet_one->Range("an1")->{Value} = 'JetDirect (9100)';
	$sheet_one->Range("ao1")->{Value} = 'SNet-Mgmt (10000)';
	$sheet_one->Range("ap1")->{Value} = 'PCAnywhere (65301)';
	$sheet_one->Range("aq1")->{Value} = 'Compaq (2301)';
	$sheet_one->Range("ar1")->{Value} = 'Other';
	$sheet_one->Range("as1")->{Value} = 'ICMP Ping';
	$sheet_one->Range("at1")->{Value} = 'Service';
	$sheet_one->Range("au1")->{Value} = 'Closed ports only';
	$sheet_one->Range("av1")->{Value} = 'DNS Resolution only';

	# SET COLOR
	$range = "a1:av1";						# cells a1-v1 will become bold, w/ blue letter and a gray shade
	$sheet_one->Range($range)->Interior->{ColorIndex} = 15;		# set background color to gray
	$sheet_one->Range($range)->Font->{FontStyle}="Bold";		# set letter style to bold
	$sheet_one->Range($range)->Font->{Size}= 11;			# set font szie to 11
	$sheet_one->Range($range)->Font->{ColorIndex}= 32;		# set letter color blue
	
	
	# ALIGN CELLS TO THE CENTER
	$sheet_one->Range("a1:aq1")->{HorizontalAlignment} = "2";		# 3 correlates to "centered"
	
	# DARKEN THE BORDERS OF THE CELLS
	$sheet_one->Range("a1:av1")->Borders->{ColorIndex} = "16";
							
}



#########################################################################	
sub process_gnmap
{
	# This function will parse and process the input file if it is specified as a grepable nmap file

	%ip_to_os;	# hash(ip_addr)-> "OS Guess"
	%ip_to_fqdn;	# hash(ip_addr)-> "Fully-qualified Domain Name"
	%ip_to_ports;	# hash(ip_addr)-> Concatenated string of open ports
	%ip_to_ping; 
	%ip_to_service;
	%ip_to_closed_ports;
	%ip_to_filtered_ports;

    $count = 0;
	print "\nProcessing Grepable Nmap File\:\n";
	
	# Process each line in the grepable nmap file
	foreach my $line (<INPUT>) {
		
		# Remove newline character at the end of the line
		chomp($line);
	
		# $count++;

		# print "$count";

		# We are only interest in "host" lines with open ports
		if ( $line !~ 'Host:' ) 
		{
			next;
			# If it isnt a host line w/ open ports, skip this line and return to begining of loop
			#if  ( $line !~ 'open' )
			#{
			#	next;
			#}
		}
		else
		{
			if ( $line =~ 'Status') 
			{
				if ($line =~ 'Up') 
				{
					@fields = split(/[\t ]/, $line);
					$current_ip = $fields[1];
					
					# Get current FQDN
					if (defined $ip_to_fqdn{$current_ip} ) 
					{
						if ( $ip_to_fqdn{$current_ip} eq "" ) 
						{
							$ip_to_fqdn{$current_ip} = $fields[2];
							$ip_to_fqdn{$current_ip} =~ s/\(//;	# get rid of the "("
							chop($ip_to_fqdn{$current_ip})	;	# get rid of the ")"
						}
					}
					else
					{
						$ip_to_fqdn{$current_ip} = $fields[2];
						$ip_to_fqdn{$current_ip} =~ s/\(//;	# get rid of the "("
						chop($ip_to_fqdn{$current_ip})	;	# get rid of the ")"
					}
					# Add port number to list of open ports for the current IP
					#if( defined ($ip_to_ports{$current_ip}) )
					#{ #append the string
        			#		$ip_to_ports{$current_ip} .= "\,ICMP";
      				#}
      				#else{ 
      				#	$ip_to_ports{$current_ip} = "ICMP";
      				#}
					if (!defined ($ip_to_ping{$current_ip}) )
					{
						# print "\tFound: $current_ip --> pingable\n"; # debugging purposes
						$ip_to_ping{$current_ip} = "ICMP";
					}
				}
			}
			else
			{		
				# Take each element of the line into an array
				@fields = split(/[\t ]/, $line);	# parse by spaces and tabs
				# Get current IP address
				$current_ip = $fields[1];
				# print "\tFound: $current_ip \n"; # debugging purposes
		
				# Get current FQDN
				if (defined $ip_to_fqdn{$current_ip} ) 
				{
					if ( $ip_to_fqdn{$current_ip} eq "" ) 
					{
						$ip_to_fqdn{$current_ip} = $fields[2];
						$ip_to_fqdn{$current_ip} =~ s/\(//;	# get rid of the "("
						chop($ip_to_fqdn{$current_ip})	;	# get rid of the ")"
					}
				}
				else
				{
					$ip_to_fqdn{$current_ip} = $fields[2];
					$ip_to_fqdn{$current_ip} =~ s/\(//;	# get rid of the "("
					chop($ip_to_fqdn{$current_ip})	;	# get rid of the ")"
				}
		
				# Process the line for Open TCp ports and OS Guess (if any)
				$os_guess_flag = 0; # We're not sure if there was an OS guess made at this point
		
				# For each element on the line
				foreach $element (@fields){
			
				# If current element is describing an open TCP port, add it to list of open tcp ports
				if ( $element =~ "open\/tcp" ) 
				{
				
					# Breaks up the element describing the open port (example: "21/open/tcp//ftp///") and stores the open port number ("21" in the example) into $temp[0]
					#@temp = &quotewords('/',0,$element);
					
					@temp = split(/[\/]/, $element);
				
					# Add port number to list of open ports for the current IP
					if( defined ($ip_to_ports{$current_ip}) ){ #append the string
        					$ip_to_ports{$current_ip} .= "\,$temp[0]";
      				}
      				else{ 
      					$ip_to_ports{$current_ip} = "$temp[0]";
      				}
					
					@buffer = split('Ports: ',$line);
					@ports = split(', ',$buffer[1]);

					$index = 0;
					foreach my $portemp (@ports) 
					{
						@port_number = split (/\//,$portemp);
						# print ("Found port$port_number[0],$temp[0], ",length($port_number[0])," ", length($temp[0]),"\n");
						if ( $temp[0] eq $port_number[0] )  
						{
							# print ("Inside Found port$port_number[0],$port_number[6]\n");
							if ( $port_number[6] ne "" )  # version found
							{
								if( defined ($ip_to_service{$current_ip}) )
								{ #append the string
        							$ip_to_service{$current_ip} .= "\,$port_number[6]";
      							}
      							else
								{
									# print "===>$port_number[6]";
      								$ip_to_service{$current_ip} = "$port_number[6]";
      							}
							}
						}
					}
				}
				else  # element was not open/tcp
				{ 
					if ( $element =~ "closed\/tcp" ) 
					{
						@temp = split(/[\/]/, $element);
						#print "Found closed=>$current_ip,$temp[0]\n";
						# Add port number to list of open ports for the current IP
						if( defined ($ip_to_closed_ports{$current_ip}) ){ #append the string
        						$ip_to_closed_ports{$current_ip} .= "\,$temp[0]";
      					}
      					else
						{ 
      						$ip_to_closed_ports{$current_ip} = "$temp[0]";
      					}	
					}
					#else
					#{
					#	if ( $element =~ "filtered\/tcp" )
					#	{
					#		@buffer = split(/[\t ]/,$line);
					#		if ( defined $ip_to_fqdn{$buffer[1]} ) 
					#		{
					#			if ( $ip_to_fqdn{$buffer[1]} ne "" )
					#			{
					#				print "Found dns name-->",$buffer[2],"\n";
					#				$buffer[2] =~ s/[\)\(]//g;
					#				$ip_to_fqdn{$buffer[1]} = $buffer[2];
					#			}
					#		}
					#	}
					#}
				} # end of else
			
				# If OS detection present, start gathering OS guesses
				if ($os_guess_flag == 1) {
				
					# Check for end of OS guesses - signified when the current element equals "Seq"
					if ($element eq "Seq") {
						# We are done gathering ports and OS detection
						last;
					}else {
						# Add element to OS guesses - concatenation
						$ip_to_os{$current_ip} .= "$element ";
					}
				}	
			
				# If the current element is 'OS:' then the next few elements to follow are the actualy OS guess
				if ($element =~ "OS\:") {
					# The next few elements are OS guesses, setting the flag to one will start collecting them
					$os_guess_flag = 1;
				}	
			}	
		} # The entire line is now processed
  	} 
    } # The entire file is now processed

	foreach $closedport_ip ( keys %ip_to_closed_ports ) 
	{
		if (defined ($ip_to_ports{$closedport_ip})) 
		{ 
			delete $ip_to_closed_ports{$closedport_ip};
		}
	}

	# foreach  $ip ( keys %ip_to_closed_ports ) { print "only_closed->$ip\n";}
 	close(INPUT);
}


########################################################################	

sub print_results() {
	# The port scan result file has now been process and the hashs %ip_to_os, %ip_to_fqdn, %ip_to_ports
	# have been populated (regardless of the portscan type).  Now take these hashes and populate the excel spreadsheet
	
	$linecount = 2; # Row Number 1 is the header
	
	
	# For each IP discovered print out (IP Address; FQDN; OS Guess; ftp..https; other)
	foreach my $ip (keys %ip_to_ports) {
		
		# Printout IP Address
		$sheet_one->Range("a$linecount")->{Value} = $ip;
		
		# Printout FQDN
		$sheet_one->Range("b$linecount")->{Value} = $ip_to_fqdn{$ip};
		
		# Printout OS Guess
		$sheet_one->Range("c$linecount")->{Value} = $ip_to_os{$ip};
		# Process and print open ports
		#@open_ports = &quotewords(',',0,$ip_to_ports{$ip});
		@open_ports = split(/[\,]/, $ip_to_ports{$ip});
		
		# Get rid of dupicate ports
		my %seen = ();
		my @uniq_ports = grep { ! $seen{$_} ++} @open_ports;
				
		# Check for common ports (ftp, telnet,...https)and places X's in spreadsheet where necessary
		@common_ports = ('21', '22', '23', '25', '53', '67', '69', '80', '111', '113', '123', '135', '139', '161', '280', '389',
			             '443', '445', '500', '1352', '1433', '1494', '1521', '1527', '1723', '1776', '2301',
			             '2701', '2967', '3306', '3389', '5060', '5560', '5631', '5800', '5900',
		                 '6000' , '9100', '10000', '65301');
		my (%union, %isect);
		foreach $i (@uniq_ports, @common_ports) { $union{$i}++ && $isect{$i}++ }
		foreach my $e (keys %isect) {
			if ($e eq '21')  { $sheet_one->Range("d$linecount")->{Value} = "X"; }
			if ($e eq '23')  { $sheet_one->Range("e$linecount")->{Value} = "X"; }
			if ($e eq '25')  { $sheet_one->Range("f$linecount")->{Value} = "X"; }
			if ($e eq '53')  { $sheet_one->Range("g$linecount")->{Value} = "X"; }
			if ($e eq '80')  { $sheet_one->Range("h$linecount")->{Value} = "X"; }
			if ($e eq '139') { $sheet_one->Range("i$linecount")->{Value} = "X"; }
			if ($e eq '443') { $sheet_one->Range("j$linecount")->{Value} = "X"; }
			if ($e eq '22') { $sheet_one->Range("k$linecount")->{Value} = "X"; }
			if ($e eq '67') { $sheet_one->Range("l$linecount")->{Value} = "X"; }
			if ($e eq '69') { $sheet_one->Range("m$linecount")->{Value} = "X"; }
			if ($e eq '111') { $sheet_one->Range("n$linecount")->{Value} = "X"; }
			if ($e eq '113') { $sheet_one->Range("o$linecount")->{Value} = "X"; }
			if ($e eq '123') { $sheet_one->Range("p$linecount")->{Value} = "X"; }
			if ($e eq '135') { $sheet_one->Range("q$linecount")->{Value} = "X"; }
			if ($e eq '161') { $sheet_one->Range("r$linecount")->{Value} = "X"; }
			if ($e eq '280') { $sheet_one->Range("s$linecount")->{Value} = "X"; }
			if ($e eq '389') { $sheet_one->Range("t$linecount")->{Value} = "X"; }
			if ($e eq '445') { $sheet_one->Range("u$linecount")->{Value} = "X"; }
			if ($e eq '500') { $sheet_one->Range("v$linecount")->{Value} = "X"; }
			if ($e eq '1352') { $sheet_one->Range("w$linecount")->{Value} = "X"; }
			if ($e eq '1433') { $sheet_one->Range("x$linecount")->{Value} = "X"; }
			if ($e eq '1494') { $sheet_one->Range("y$linecount")->{Value} = "X"; }
			if ($e eq '1521') { $sheet_one->Range("z$linecount")->{Value} = "X"; }
			if ($e eq '1527') { $sheet_one->Range("aa$linecount")->{Value} = "X"; }
			if ($e eq '1723') { $sheet_one->Range("ab$linecount")->{Value} = "X"; }
			if ($e eq '1776') { $sheet_one->Range("ac$linecount")->{Value} = "X"; }
			if ($e eq '2701') { $sheet_one->Range("ad$linecount")->{Value} = "X"; }
			if ($e eq '2967') { $sheet_one->Range("ae$linecount")->{Value} = "X"; }
			if ($e eq '3306') { $sheet_one->Range("af$linecount")->{Value} = "X"; }
			if ($e eq '3389') { $sheet_one->Range("ag$linecount")->{Value} = "X"; }
			if ($e eq '5060') { $sheet_one->Range("ah$linecount")->{Value} = "X"; }
			if ($e eq '5560') { $sheet_one->Range("ai$linecount")->{Value} = "X"; }
			if ($e eq '5631') { $sheet_one->Range("aj$linecount")->{Value} = "X"; }
			if ($e eq '5800') { $sheet_one->Range("ak$linecount")->{Value} = "X"; }
			if ($e eq '5900') { $sheet_one->Range("al$linecount")->{Value} = "X"; }
			if ($e eq '6000') { $sheet_one->Range("am$linecount")->{Value} = "X"; }
			if ($e eq '9100') { $sheet_one->Range("an$linecount")->{Value} = "X"; }
			if ($e eq '10000') { $sheet_one->Range("ao$linecount")->{Value} = "X"; }
			if ($e eq '65301') { $sheet_one->Range("ap$linecount")->{Value} = "X"; }
			if ($e eq '2301') { $sheet_one->Range("aq$linecount")->{Value} = "X"; }
			#if ($e eq 'ICMP') { $sheet_one->Range("ar$linecount")->{Value}= "X"; }

		}
		
		# Remove common ports from general list (to generate "Other" list)
		my (%seen2, $other);
		@seen2 {@uniq_ports} = ();
		delete @seen2 {@common_ports};
		my @others = keys %seen2;			# These are the 'Other' ports
		my @sorted_ports = sort {$a <=> $b} @others;	# sort ports numerically
		
		# Turns sorted 'others' array into a comma seperated string called $other
		foreach my $p (@sorted_ports) {
			$other .= "$p, ";
		}
		chop ($other); # get rid of last space
		chop ($other); # get rid of last comma
		
		# Insert other ports into the Other column of the spreadsheet
		$sheet_one->Range("ar$linecount")->{Value} = $other;
		$sheet_one->Range("as$linecount")->{Value} = $ip_to_ping{$ip};
		$sheet_one->Range("at$linecount")->{Value} = $ip_to_service{$ip};

		# Go to next line of the spreadsheet
		 $linecount++;
	}

	foreach my $test ( keys %ip_to_ping ) 
	{
		if ( defined ($ip_to_ports{$test}) ) 
		{
		}
		else
		{
			print "Found: $test\n";
			$sheet_one->Range("a$linecount")->{Value} = $test;
			$sheet_one->Range("b$linecount")->{Value} = $ip_to_fqdn{$test};
			$sheet_one->Range("as$linecount")->{Value} = "ICMP";
			$linecount++;
		}
	}

	foreach my $test ( keys %ip_to_closed_ports ) 
	{
		if ( defined ($ip_to_ports{$test}) ) 
		{
		}
		else
		{
			if ( defined ($ip_to_ping{$test}) )
			{
			}
			else
			{
				print "Found: $test\n";
				$sheet_one->Range("a$linecount")->{Value} = $test;
				$sheet_one->Range("b$linecount")->{Value} = $ip_to_fqdn{$test};
				$sheet_one->Range("au$linecount")->{Value} = $ip_to_closed_ports{$test};
				$linecount++;
			}
		}
	}

	foreach my $dnsip ( keys %ip_to_fqdn ) 
	{
		if (defined ( $ip_to_ports{$dnsip}) ) 
		{ # nothing to do, open port found with this DNS
		}
		else
		{
			if ( defined ($ip_to_closed_ports{$dnsip}) )
			{ # closed port found
			}
			else
			{
				if( defined ($ip_to_ping{$dnsip}) )
				{ # ping entry found
				}
				else
				{ # neither ping, nor closed port, nor open port ... hence only filtered
					# print "dns only->$dnsip,",$ip_to_fqdn{$dnsip},"\n";
					if ( $ip_to_fqdn{$dnsip} ne "" )
					{
						print "Found : $dnsip\n";
						$sheet_one->Range("a$linecount")->{Value} = $dnsip;
						$sheet_one->Range("b$linecount")->{Value} = $ip_to_fqdn{$dnsip};
						$sheet_one->Range("av$linecount")->{Value} = "X";
						$linecount++;
					}

				}
			}
		}
	}
}	
	
#########################################################################
sub finish_excel() {


	# Change Cell Orientation
	$sheet_one->Range("d1:av1")->{Orientation} = 90;
	

	# Autofit the columns
	$sheet_one->Columns("a")->{ColumnWidth} = 20.5;
	$sheet_one->Columns("b")->{ColumnWidth} = 25.9;
	$sheet_one->Columns("c")->{ColumnWidth} = 26.38;
	$sheet_one->Columns("d")->AutoFit();
	$sheet_one->Columns("e")->AutoFit();
	$sheet_one->Columns("f")->AutoFit();
	$sheet_one->Columns("g")->AutoFit();
	$sheet_one->Columns("h")->AutoFit();
	$sheet_one->Columns("i")->AutoFit();
	$sheet_one->Columns("j")->AutoFit();
	$sheet_one->Columns("k")->AutoFit();
	$sheet_one->Columns("l")->AutoFit();
	$sheet_one->Columns("m")->AutoFit();
	$sheet_one->Columns("n")->AutoFit();
	$sheet_one->Columns("o")->AutoFit();
	$sheet_one->Columns("p")->AutoFit();
	$sheet_one->Columns("q")->AutoFit();
	$sheet_one->Columns("r")->AutoFit();
	$sheet_one->Columns("s")->AutoFit();
	$sheet_one->Columns("t")->AutoFit();
	$sheet_one->Columns("u")->AutoFit();
	$sheet_one->Columns("v")->AutoFit();
	$sheet_one->Columns("w")->AutoFit();
	$sheet_one->Columns("x")->AutoFit();
	$sheet_one->Columns("y")->AutoFit();
	$sheet_one->Columns("z")->AutoFit();
	$sheet_one->Columns("aa")->AutoFit();
	$sheet_one->Columns("ab")->AutoFit();
	$sheet_one->Columns("ac")->AutoFit();
	$sheet_one->Columns("ad")->AutoFit();
	$sheet_one->Columns("ae")->AutoFit();
	$sheet_one->Columns("af")->AutoFit();
	$sheet_one->Columns("ag")->AutoFit();
	$sheet_one->Columns("ah")->AutoFit();
	$sheet_one->Columns("ai")->AutoFit();
	$sheet_one->Columns("aj")->AutoFit();
	$sheet_one->Columns("ak")->AutoFit();
	$sheet_one->Columns("al")->AutoFit();
	$sheet_one->Columns("am")->AutoFit();
	$sheet_one->Columns("an")->AutoFit();
	$sheet_one->Columns("ao")->AutoFit();
	$sheet_one->Columns("ap")->AutoFit();
	$sheet_one->Columns("aq")->AutoFit();
	$sheet_one->Columns("ar")->{ColumnWidth} = 21.5;
	$sheet_one->Columns("as")->AutoFit();
	$sheet_one->Columns("at")->{ColumnWidth} = 21.5;
	$sheet_one->Columns("au")->{ColumnWidth} = 21.5;
	$sheet_one->Columns("av")->AutoFit();

	# WrapText the "OS Guess" and "Other ports" columns
	$sheet_one->Range("c2:c$linecount")->{WrapText} = 1;
	$sheet_one->Range("ar2:ar$linecount")->{WrapText} = 1;
	$sheet_one->Range("at2:au$linecount")->{WrapText} = 1;
	$sheet_one->Range("au2:au$linecount")->{WrapText} = 1;

	# Center the X's for commons ports (i.e. ftp, telnet,..,https)
	$sheet_one->Range("d2:aq$linecount")->{HorizontalAlignment} = "3";
	$sheet_one->Range("av2:av$linecount")->{HorizontalAlignment} = "3";
	$sheet_one->Range("as2:as$linecount")->{HorizontalAlignment} = "3";

	# Align 'Other' column cells to bottom-left
	$sheet_one->Range("ar2:ar$linecount")->{HorizontalAlignment} = "2";
	$sheet_one->Range("ar2:ar$linecount")->{VerticalAlignment} = "3";
	
	# Sort the spreadsheet by the IP Address column
	$sheet_one->Range("a2:av$linecount")->Sort("c1",1);

	# Save and close the new excel file
	$new_workbook->SaveAs($excel_doc);
	$new_workbook->Close;

}
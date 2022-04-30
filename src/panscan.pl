# -- Scan a directory tree for PAN in DOC, XLS, TXT and MSG
# -- 

# -------------------------------------------------------------
# -- Get current working directory
# -------------------------------------------------------------
use Cwd;
my $path = getcwd;
$path =~ s/\//\\\\/gsx;

# -------------------------------------------------------------
# -- Let's check if catdoc is installed
# -------------------------------------------------------------
unless (-e $path.'\\catdoc.exe')
{
  print "Error | ".$path."\\catdoc.exe is not present.\nAborting.\n";
  exit;
}

unless (-e $path.'\\xls2csv.exe')
{
  print "Error | ".$path."\\catxls.exe is not present.\nAborting.\n";
  exit;
}

# -------------------------------------------------------------
# -- Let's check if strings is installed
# -------------------------------------------------------------
unless (-e $path.'\\strings.exe')
{
  print "Error | ".$path."\\strings.exe is not present.\nAborting.\n";
  exit;
}

# -------------------------------------------------------------
# -- Starting the search
# -------------------------------------------------------------
if (@ARGV)
{
	use File::Find;
	print "Searching for txt, doc, msg and xls files.\nIt can take a while ...\nOnly printing suspicious files.\n\n";
}
else
{
	print "Error | Syntax: panscan.exe FolderName.\nAborting\n";
	exit;
}


# -------------------------------------------------------------
# -- Getting files path
# -------------------------------------------------------------
find sub { 
	if( grep(/\.(xls)$/,$File::Find::name) )
	{
		push(@xls, $File::Find::name);
	}
	elsif( grep(/\.(txt)$/,$File::Find::name) )
	{
		push(@txt, $File::Find::name);
	}
	elsif( grep(/\.(doc)$/,$File::Find::name) )
	{
		push(@doc, $File::Find::name);
	}	
	elsif( grep(/\.(msg)$/,$File::Find::name) )
	{
		push(@msg, $File::Find::name);
	}
}, @ARGV;

# -------------------------------------------------------------
# -- Processing Excel files
# -------------------------------------------------------------
foreach $file (@xls) {
   #print "\n -- Processing : ".$file." --\n";
   $result = `$path\\xls2csv.exe "$file"`;
   if (search_digits($result) == 1)
   {
		print $file."\n";
   }
   # else
   # {
		# print "Clean file !\n";
   # }
}

# -------------------------------------------------------------
# -- Processing Word files
# -------------------------------------------------------------
foreach $file (@doc) {
   #print "\n -- Processing : ".$file." --\n";
   $result = `$path\\catdoc.exe -a "$file"`;
   if (search_digits($result) == 1)
   {
		print $file."\n";
   }
   # else
   # {
		# print "Clean file !\n";
   # }
}

# -------------------------------------------------------------
# -- Processing Text files
# -------------------------------------------------------------
foreach $file (@txt) {
   #print "\n -- Processing : ".$file." --\n";
   $file =~ s/\//\\/gsx;
   $result = `type "$file"`;
   if (search_digits($result) == 1)
   {
		print $file."\n";
   }
   # else
   # {
		# print "Clean file !\n";
   # }
}

# -------------------------------------------------------------
# -- Processing Outlook msg files
# -------------------------------------------------------------
foreach $file (@msg) {
   #print "\n -- Processing : ".$file." --\n";
   $file =~ s/\//\\/gsx;
   $result = `$path\\strings.exe "$file"`;
   if (search_digits($result) == 1)
   {
		print $file."\n";
   }
   # else
   # {
		# print "Clean file !\n";
   # }
}

print "\nSearch completed.";

# -------------------------------------------------------------
# -- Function searching for digit sequences
# -------------------------------------------------------------
sub search_digits {
	@tmp = split(/\n/, $_[0]);
	foreach $CB (@tmp)
	{ 
		if ( grep(/((\d{4}[\s-]*){3}\d{4})|\d{16}/, $CB) )
		{
			if (&luhn_algorithn() == 1)
			{
				return (1); 
			}
		}
	}
	return (0);
}

# -------------------------------------------------------------
# -- Check if the PAN is a valid one
# -------------------------------------------------------------
sub luhn_algorithn()
{
    my ($sum,$odd, $number);
    $number = $CB;

    foreach my $n (reverse split(//, $number))
    {
        $odd =! $odd;
        if ($odd)
        {
            $sum += $n;
        }
        else
        {
            my $x = 2 * $n;
            $sum += ($x > 9) ? ($x - 9) : $x;
        }
    }
    return (($sum % 10) == 0);
}
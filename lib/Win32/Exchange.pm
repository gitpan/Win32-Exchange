# Win32::Exchange
# Freely Distribute the code without modification.
#
# Creates and Modifies Exchange 5.5 and 2K Mailboxes
# (eventually it will do more, but for now, that's the scope)
#
# This is the culmination of 3 years of work in building Exchange Mailboxes, but now as a module.
# It uses Win32::OLE exclusively (and technically is just a wrapper for the underlying OLE calls).
# 
# This build is tested and works with ActivePerl build 633 (and Win32::OLE .1502)
# There is not currently a package that is tested on older (non-multi-threading versions of
# ActivePerl)... My guess is that it may work except for the SetPerms and SetOwner subs but,
# remember..  That's a guess.
# 
# Sorry... :(
#

package Win32::Exchange;

use strict;
use vars qw ($VERSION $Version $DEBUG);

use Win32::OLE;
Win32::OLE->Initialize(Win32::OLE::COINIT_OLEINITIALIZE);
use Win32::Exchange::Const;

Win32::OLE->Option('_Unique' => 1);
#@ISA = qw(Win32::OLE);

my $Version;
my $VERSION = $Version = "0.032";
my $DEBUG = 1;


sub new {
  my $server;
  my $ver = "";
  if (scalar(@_) == 1) {
    if ($_[0] eq "5.5" || $_[0] eq "6.0") {
      $ver = $_[0];
    } else {
      $server = $_[0];
    }
  } elsif (scalar(@_) == 2) {
    if ($_[0] eq "Win32::Exchange") {
      if ($_[1] eq "5.5" || $_[1] eq "6.0") {
        $ver = $_[1];
      } else {
        $server = $_[1];
      }

    } else {
      _ReportArgError("new",scalar(@_));
    }
  } else {
    _ReportArgError("new",scalar(@_));
    return 0;
  }

  my $class = "Win32::Exchange";
  my $ldap_provider = {};

  if ($ver eq "") {
    my %version;
    if (!Win32::Exchange::GetVersion($server,\%version)) {
      return undef;
    } else {
      $ver = $version{'ver'}
    }
  }
  if ($ver eq "5.5") {
    #Exchange 5.5
    if ($ldap_provider = Win32::OLE->new('ADsNamespaces')) {
      return bless $ldap_provider,$class;
    } else {
      _DebugComment("Failed creating ADsNamespaces object\n",1);
      return undef;
    }
  } elsif ($ver eq "6.0") {
    #Exchange 2000
    if ($ldap_provider = Win32::OLE->new('CDO.Person')) {
      return bless $ldap_provider,$class;
    } else {
      _DebugComment("Failed creating CDO.Person object\n",1);
      return undef;
    }
  } else {
    _DebugComment("Unable to verify version information for version: $ver\n",1);
    return undef;
  }
}

sub DESTROY {
  my $object = shift;
  bless $object,"Win32::OLE";
  return undef;
}

sub GetLDAPPath {
  my $ldap_provider;
  my $server_name;
  my $ldap_path;
  my $return_point;
  if (scalar(@_) == 3) {
    $server_name = $_[0];
    $ldap_path = "LDAP://$server_name";
    $return_point = 1;
  } elsif (scalar(@_) == 4) {
    $ldap_provider = $_[0];
    $server_name = $_[1];
    $return_point = 2;
  } else {
    _ReportArgError("GetLDAPPath",scalar(@_));
    return 0;
  }
  my $result;
  if (_AdodbExtendedSearch($server_name,"LDAP://$server_name","(objectClass=Computer)","rdn,distinguishedName",$result)) {
    _DebugComment("result = $result\n",2);
    if ($result =~ /cn=.*,cn=Servers,cn=Configuration,ou=(.*),o=(.*)/) {
      my $returned_ou = $1;
      my $returned_o = $2;
      $_[$return_point]=$returned_o;
      $_[($return_point+1)]=$returned_ou;
      _DebugComment("ou=$returned_ou\no=$returned_o\n",2);
      return 1;
    } else {
      _DebugComment("result = $result\n",2);
      _DebugComment("result from ADODB search failed to produce an acceptable match\n",1);
      return 0;
    }
  } else {
    _DebugComment("ADODB search failed\n",1);
    return 0;  
  }
}

sub GetVersion {
  my $server_name;
  my $error_num;
  my $error_name;
  if (scalar(@_) == 2) {
    $server_name = $_[0];
  } elsif (scalar(@_) == 3) {
    if ($_[0] eq "Win32::Exchange") {
      $server_name = $_[1];
    } else {
      _ReportArgError("GetVersion",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("GetVersion",scalar(@_));
    return 0;
  }
  my $original_ole_warn_value = $Win32::OLE::Warn;
  $Win32::OLE::Warn = 0;
  my $serial_val;
  my $serial_version_check_obj = Win32::OLE->new('CDOEXM.ExchangeServer'); #substantiates the possible existance of e2k
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    if ($error_num eq "0x80040154" ||
        $error_num eq "0x800401f3") {
      #0x80040154 Class not registered
      #0x800401f3 Invalid class string
      _DebugComment("The Exchange 2000 client tools don't look to be installed on this machine\n",2);
      if (!_E55VersionInfo($server_name,$serial_val)) {
        _DebugComment("Error getting version information from Exchange 5.5\n",1);
        $Win32::OLE::Warn = $original_ole_warn_value;
        return 0;
      }
    } else {
      _DebugComment("error: $error_num - $error_name on $server_name encountered while trying to perform GetVersion\n",1);
      $Win32::OLE::Warn = $original_ole_warn_value;
      return 0;
    }
  } else {
    _DebugComment("found e2k tools, so we'll look and see what version of Exchange you have.\n",3);
    if (!_E2kVersionInfo($server_name,$serial_val)) {
      _DebugComment("Error getting version information from Exchange 2000 tools, let's try the Exch 5.5 way\n",3);
      if (!_E55VersionInfo($server_name,$serial_val)) {
        _DebugComment("Error getting version information trying the Exch 5.5 way\n",3);
        _DebugComment("Error getting version information\n",1);
        $Win32::OLE::Warn = $original_ole_warn_value;
        return 0;
      }
    }
  }
  $Win32::OLE::Warn = $original_ole_warn_value; 

  if ($serial_val =~ /Version (.*) \(Build (.*): Service Pack (.*)\)/i) {
    my %return_struct;
    $return_struct{ver}= $1;
    $return_struct{build}= $2;
    $return_struct{sp}= $3;
    if ($return_struct{sp} < 2 && $return_struct{ver} eq "6.0") {
      _DebugComment("It's possible that some of the E2K permissions functions will fail due to an incompatible E2K Service Pack level (please see the HTML docs for details)\n",2)
    }
    if (scalar(@_) == 2) {
      %{$_[1]} = %return_struct;
    } else {
      %{$_[2]} = %return_struct;
    }
    return 1;
  } else {
    return 0;
  }
}
 
sub _E55VersionInfo {
  my $server_name;
  my $error_num;
  my $error_name;
  if (scalar(@_) == 2) {
    $server_name = uc($_[0]);
  } else {
    _ReportArgError("_E55VersionInfo",scalar(@_));
    return 0;
  }
  my $serial_val;
  my $provider;
  my $org;
  my $ou;
  $provider = Win32::Exchange->new("5.5");
  if (!$provider) {
    _DebugComment("new provider create in GetVersion (E55) failed\n",1);
    return 0;
  }
  if ($provider->GetLDAPPath($server_name,$org,$ou)) {
    _DebugComment("returned -> o=$org,ou=$ou\n",3);
  } else {
    _DebugComment("Error Returning from GetLDAPPath in GetVersion (E55)\n",1);
    return 0;
  }
  bless $provider,"Win32::OLE";
  my $exch_server_obj = $provider->GetObject("","LDAP://$server_name/cn=$server_name,cn=Servers,cn=Configuration,ou=$ou,o=$org");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed getting the server object in Server container for version info (E55): $error_num,$error_name\n",1);
    return 0;
  }
  $exch_server_obj->GetInfoEx(['serialNumber'],0);
  $serial_val = $exch_server_obj->{"serialNumber"};
  if ($serial_val =~ /Version (.*) \(Build (.*): Service Pack (.*)\)/i) {
    $_[1] = $serial_val;
    return 1;
  } else {
    _DebugComment("GetVersion failed to produce acceptable results (E55)\n",1);
    return 0;
  }
}
 
sub _E2kVersionInfo {
  my $error_num;
  my $error_name;
  if (scalar(@_) != 2) {
    _ReportArgError("_E2kVersionInfo",scalar(@_));
  }
  my $server_name = $_[0];
  my $exchange_server = Win32::OLE->new("CDOEXM.ExchangeServer");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Failed creating object for version information (E2K) on $server_name -> $error_num ($error_name)\n",1);
      return 0;
  }
  $exchange_server->DataSource->Open($server_name);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    if ($error_num eq "0x80072032") {
      #This error might be there if the server is on another domain...  not sure..  I'll need to research more.
      #It happenned on an E5.5 server anyway so I didn't need the E2K version strucure.
      _DebugComment("Failed opening object for version information (E2K) on $server_name -> $error_num ($error_name)\n",2);
    } else {
      _DebugComment("Failed opening object for version information (E2K) on $server_name -> $error_num ($error_name)\n",1);
    }
    return 0;
  }
 
  #example output:
  #Version 5.5 (Build 2653.23: Service Pack 4)
  #Version 6.0 (Build 6249.4: Service Pack 3)
 
  if ($exchange_server->{ExchangeVersion} ne "") {
    $_[1] = $exchange_server->{ExchangeVersion};
    return 1;
  } else {
    _DebugComment("Failed failed to produce valid version info for $server_name\n",1);
    return 0;
  }
}

sub AdodbSearch {
  my $x;
  my $server_name;
  my $filter;
  my $columns;
  my $error_num;
  my $error_name;
  my $return_point;
  my $ldap_path;
  if (scalar(@_) > 3) {
      $server_name = $_[0];
      $filter = $_[1];
      $columns = $_[2];
    if (scalar(@_) == 4) {
      $ldap_path = "LDAP://$server_name";
      $return_point=3;
    } elsif (scalar(@_) == 5) {
      $return_point=4;
      $ldap_path = $_[3];
    } else {
      _ReportArgError("AdodbSearch",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("AdodbSearch",scalar(@_));
    return 0;
  }
  my @ado_columns = split (/,/,$columns);
  my $Conn = Win32::OLE->new("ADODB.Connection");
  $Conn->{'Provider'} = "ADsDSOObject";
  $Conn->Win32::OLE::Open("Active Directory Provider;UID=;PWD=");
  my $path = "<$ldap_path>;$filter;$columns;subtree";
  my $RS = $Conn->Win32::OLE::Execute($path);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("path=$ldap_path\nfilter=$filter\ncolumns=$columns\n",2);
    _DebugComment("Failed Executing ADODB Execute command on $server_name -> $error_num ($error_name)\n",1);
    return 0;
  }
  if ($RS->RecordCount == 0) {
    _DebugComment("path=$ldap_path\nfilter=$filter\ncolumns=$columns\n",2);
    _DebugComment("AdodbSearch yeilded no results for search on $server_name -> $error_num ($error_name)\n",1);
  } elsif ($RS->RecordCount > 1) {
    _DebugComment("path=$ldap_path\nfilter=$filter\ncolumns=$columns\n",2);
    _DebugComment("AdodbSearch yeilded more than 1 result for search on $server_name -> $error_num ($error_name)\n",1);
    return 0;
  } else {
    $_[$return_point] = $RS->Fields($ado_columns[1])->value 
  }
}

sub _AdodbExtendedSearch {
  my $server_name;
  my $path;
  my $filter;
  my $columns;
  my $error_num;
  my $error_name;
  my $fuzzy;
  my $return_point;
  if (scalar(@_) > 4) {
    $server_name = $_[0];
    $path = $_[1];
    $filter = $_[2];
    $columns = $_[3];
    if (scalar(@_) == 5) {
      $return_point = 4;
    } elsif (scalar(@_) == 6) {
      $fuzzy = $_[4];
      $return_point = 5;
    }
  } else {
    _ReportArgError("_AdodbExtendedSearch (".scalar(@_));
    return 0;
  }
  my @cols = split (/,/,$columns);
  if (scalar(@cols) != 2) {
    _DebugComment("Only 2 columns can be sent to _AdodbExtendedSearch (total recieved = ".scalar(@cols).")\n",1);
  }
  my $option;
  if ($path =~ /^LDAP:\/\/RootDSE\/(.*)/i) {
    $option = $1;
    my $RootDSE = Win32::OLE->GetObject("LDAP://RootDSE");
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Failed creating object for _AdodbExtendedSearch on $server_name -> $error_num ($error_name)\n",1);
      return 0;
    }
    my $actual_ldap_path = $RootDSE->Get($option);
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Failed creating object for _AdodbExtendedSearch on $server_name -> $error_num ($error_name)\n",1);
      return 0;
    }
    $path = "LDAP://".$actual_ldap_path;
  }
  my $string = "<$path>;$filter;$columns;subtree";
  my $Com = Win32::OLE->new("ADODB.Command");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("path=$path\nfilter=$filter\ncolumns=$columns\n",2);
      _DebugComment("Failed creating ADODB.Command object for _AdodbExtendedSearch on $server_name -> $error_num ($error_name)\n",1);
      return 0;
  }
  my $Conn = Win32::OLE->new("ADODB.Connection");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("path=$path\nfilter=$filter\ncolumns=$columns\n",2);
      _DebugComment("Failed creating ADODB.Connection object for version information (E55) on $server_name -> $error_num ($error_name)\n",1);
      return 0;
  }
  $Conn->{'Provider'} = "ADsDSOObject";
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("path=$path\nfilter=$filter\ncolumns=$columns\n",2);
      _DebugComment("Failed executing ADODB.Command for version information (E55) on $server_name -> $error_num ($error_name)\n",1);
      return 0;
  }
  $Conn->{Open} = "Win32-Exchange a perl module";
  $Com->{ActiveConnection} = $Conn;
  $Com->{CommandText} = $string;
  $Com->{Properties}->{"Page Size"} = 99; #One less than the default of 100 for Exchange so we don't return an empty resultset if more than 100 results are found
  my $RS = $Com->Execute();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("path=$path\nfilter=$filter\ncolumns=$columns\n",2);
      _DebugComment("Failed executing ADODB.Command for version information (E55) on $server_name -> $error_num ($error_name)\n",1);
      return 0;
  }
  my $not_found = 1;
  my $search_val = "";
  while ($search_val eq "") {
    if ($fuzzy != 0) {
      _DebugComment("fuzzy=$fuzzy\n",3);
      if ($RS->Fields($cols[($fuzzy - 1)])->value =~ /$server_name/i) {
        if (ref($RS->Fields($cols[($fuzzy - 1)])->value) eq "ARRAY") {
          _DebugComment("found ".@{$RS->Fields($cols[1])->value}[0]."\n",3);
          $search_val = @{$RS->Fields($cols[1])->value}[0]; 
          $_[$return_point] = $search_val;
          return 1;
        } else {
          _DebugComment("not found ".$RS->Fields($cols[1])->value."\n",3);
          $search_val = $RS->Fields($cols[1])->value; 
          $_[$return_point] = $search_val;
          return 1;
        }
      }
    } else {
      if (lc($server_name) eq lc($RS->Fields($cols[0])->value)) {
        if (ref($RS->Fields($cols[1])->value) eq "ARRAY") {
          _DebugComment("found (not fuzzy) (ARRAY)".$RS->Fields($cols[1])->value."\n",3);
          $search_val = @{$RS->Fields($cols[1])->value}[0]; 
          $_[$return_point] = $search_val;
          return 1;
        } else {
          _DebugComment("found (not fuzzy) (string)".$RS->Fields($cols[1])->value."\n",3);
          $search_val = $RS->Fields($cols[1])->value; 
          $_[$return_point] = $search_val;
          return 1;
        }
      }
    }
    _DebugComment($RS->Fields($cols[0])->value." - ".$RS->Fields($cols[1])->value."\n",3);
    if ($RS->EOF) {
      $search_val = "-1";
    }
    $RS->MoveNext;
  }
  if ($search_val eq "-1") {
    _DebugComment("Unable to match valid data for your search on $server_name\n",1);
    return 0;
  }
}

sub LocateMailboxStore {
  my $store_server;
  my $storage_group;
  my $mb_store;
  my $count = "no";
  if ($_[0] eq "Win32::Exchange") {
    if (scalar(@_) > 4) {
      if (scalar(@_) == 5) {
      } elsif (scalar(@_) == 6) {
        if (ref($_[5]) eq "ARRAY") {
          $count = "yes";
        } else {
          _DebugComment("the fifth argument passed to LocateMailboxStore must be an array (but is optional).\n",1);
          return 0;  
        }
      } else {
        _ReportArgError("LocateMailboxStore [E2K] (".scalar(@_));
       return 0;  
       }
    } else {
      _ReportArgError("LocateMailboxStore [E2K] (".scalar(@_));
      return 0;  
    }
  } else {
    if (scalar(@_) > 3) {
      if (scalar(@_) == 4) {
      } elsif (scalar(@_) == 5) {
        if (ref($_[4]) eq "ARRAY") {
          $count = "yes";
        } else {
          _DebugComment("the fifth argument passed to LocateMailboxStore must be an array (but is optional).\n",1);
          return 0;  
        }
      } else {
        _ReportArgError("LocateMailboxStore [E2K] (".scalar(@_));
       return 0;  
       }
    } else {
      _ReportArgError("LocateMailboxStore [E2K] (".scalar(@_));
      return 0;  
    }
  }
  
  my $ldap_path;
  my $mb_count;
  my %storage_groups;
  $store_server = $_[0];
  $storage_group = $_[1];
  $mb_store = $_[2];
  if (_EnumStorageGroups($store_server,\%storage_groups)) {
    if ($count eq "yes") {
      foreach my $sg (keys %storage_groups) {
        $mb_count += scalar(keys %{$storage_groups{$sg}}); 
      }
      push (@{$_[4]},scalar(keys %storage_groups)); 
      push (@{$_[4]},$mb_count); 
    }
    if (_TraverseStorageGroups(\%storage_groups,$store_server,$storage_group,$mb_store,$ldap_path)) {
      $_[3] = $ldap_path;
      return 1;
    } else {
      _DebugComment("Unable to locate valid mailbox store for mailbox creation.\n",1);
      return 0;          
    }
  } else {
    _DebugComment("Unable to locate valid storage group for mailbox creation.\n",1);
    return 0;          
  }
}

sub _EnumStorageGroups {
  my $server_name;
  my $error_num;
  my $error_name;
  if (scalar(@_) == 2) {
    $server_name = $_[0];
  } else {
    _ReportArgError("_EnumStorageGroups (".scalar(@_));
    return 0;
  }
  my $exchange_server = Win32::OLE->new("CDOEXM.ExchangeServer");

  $exchange_server->DataSource->Open($server_name);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening ADODB ExchangeServer object for Storage Group enumeration on $server_name -> $error_num ($error_name)\n",1);
    return 0;
  }

  my @storegroups = Win32::OLE::in($exchange_server->StorageGroups);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed enumerating Storage Groups on $server_name -> $error_num ($error_name)\n",1);
    return 0;
  }
  my %storage_groups;
  my $stor_group_obj = Win32::OLE->new("CDOEXM.StorageGroup");
  my $mbx_store_obj = Win32::OLE->new("CDOEXM.MailboxStoreDB");
  foreach my $storegroup (@storegroups) {
    $stor_group_obj->DataSource->Open($storegroup);
    _DebugComment("Stor Name = ".$stor_group_obj->{Name}."\n",3);
    foreach my $mbx_store (Win32::OLE::in($stor_group_obj->{MailboxStoreDBs})) {
      $mbx_store_obj->DataSource->Open($mbx_store);
      _DebugComment("  Mailbox Store = $mbx_store_obj->{Name}\n",3);
      $storage_groups{$stor_group_obj->{Name}}{$mbx_store_obj->{Name}}=$mbx_store;
    }
  }

  %{$_[1]} = %storage_groups;
  return 1;
}

sub _TraverseStorageGroups {
  if (scalar(@_) != 5) {
    _ReportArgError("_TraverseStorageGroups [E2K] (".scalar(@_));
    return 0;
  }
  if (ref($_[0]) ne "HASH") {
    _DebugComment("Storage group object is not a hash\n",1);
    return 0;
  }
  my %storage_groups = %{$_[0]};
  my $info_store_server = $_[1];
  my $storage_group = $_[2];
  my $mb_store = $_[3];
  my $ldap_path;
  if (scalar(keys %storage_groups) == 0) {
      _DebugComment("No Storage Groups were found\n",1);
      return 0;
  }
  my $sg;
  my $mb;
  foreach $sg (keys %storage_groups) {
    if (scalar(keys %storage_groups) == 1) {
      foreach $mb (keys %{$storage_groups{$sg}}) {
        if (scalar(keys %{$storage_groups{$sg}}) == 1 || $mb eq $mb_store && $mb_store ne "") {
          $_[4] = "LDAP://$info_store_server/".$storage_groups{$sg}{$mb}; 
          return 1;
        } else {
          next;
        }
      }
      _DebugComment("Error locating proper storage group and mailbox db for mailbox creation (1SG)\n",1);
      return 0;
    } elsif ($sg eq $storage_group && $storage_group ne "") {
      foreach $mb (keys %{$storage_groups{$sg}}) {
        if (scalar(keys %{$storage_groups{$sg}}) == 1 || $mb eq $mb_store && $mb_store ne "") {
          $_[4] = "LDAP://$info_store_server/".$storage_groups{$sg}{$mb}; 
          return 1;
        } else {
          next;
        }
      }
      _DebugComment("Error locating proper storage group and mailbox db for mailbox creation (2+SG)\n",1);
      return 0;
    }
  }
}

sub CreateMailbox {
  my $error_num;
  my $error_name;
  my $mbx;
  my $provider = $_[0];

  bless $provider,"Win32::OLE";
  
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object for Exchange Server Determination for CreateMailbox ($error_num)\n",1);
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
  if ($type eq "IPerson") {
    #IPerson returns for CDO.Person (E2K)
    if ($mbx = _E2KCreateMailbox(@_)) {
      bless $provider,"Win32::Exchange";
      bless $mbx,"Win32::Exchange";
      return $mbx;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($mbx = _E55CreateMailbox(@_)) {
      bless $provider,"Win32::Exchange";
      bless $mbx,"Win32::Exchange";
      return $mbx;
    }
  }
  bless $provider,"Win32::Exchange";
  return 0;
}

sub _E55CreateMailbox {
  my $ldap_provider;
  my $information_store_server;
  my $mailbox_alias_name;
  my $org;
  my $ou;
  my $error_num;
  my $error_name;
  my $container = "";
  my $recipients_path;
  if (scalar(@_) > 2) {
    $ldap_provider = $_[0];
    $information_store_server = $_[1];
    $mailbox_alias_name = $_[2];
    if (scalar(@_) == 3) {
      if ($ldap_provider->GetLDAPPath($information_store_server,$org,$ou)) {
        _DebugComment("returned -> o=$org,ou=$ou\n",3);
      } else {
        _DebugComment("Error Returning from GetLDAPPath\n",1);
        return 0;
      }
    } elsif (scalar(@_) == 4) {
      $container = $_[4];
    } elsif (scalar(@_) == 5) {
      $org = $_[3];
      $ou = $_[4];
    } else {
      _ReportArgError("CreateMailbox [5.5] (".scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("CreateMailbox [5.5] (".scalar(@_));
    return 0;
  }
  if ($container ne "") {
    $recipients_path = "LDAP://$information_store_server/$container";
  } else {
    $recipients_path = "LDAP://$information_store_server/cn=Recipients,ou=$ou,o=$org";
  }
  _DebugComment("path to create mailbox in: $recipients_path\n",3);
  bless $ldap_provider,"Win32::OLE";

  my $original_ole_warn_value = $Win32::OLE::Warn;
  $Win32::OLE::Warn = 0; #Turn STDERR warnings off because we probably are going to get an error (0x80072030)

  my $Recipients = $ldap_provider->GetObject("",$recipients_path);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening recipients path ($recipients_path)\nError: $error_num ($error_name)\n",1);
    return 0;
  }

  $Recipients->GetObject("organizationalPerson", "cn=$mailbox_alias_name");
  if (!ErrorCheck("0x80072030",$error_num,$error_name)) {
    if ($error_num eq "0x00000000") {
      _DebugComment("$error_num - Mailbox already exists on $information_store_server\n",1);
      $Win32::OLE::Warn=$original_ole_warn_value;
      return 0;
    } else {
      _DebugComment("Unable to lookup object $mailbox_alias_name on $information_store_server ($error_num)\n",1);
      $Win32::OLE::Warn=$original_ole_warn_value;
      return 0;
    }
  }
  _DebugComment("    Box Does Not Exist (This is good)\n",3);
  bless $ldap_provider,"Win32::Exchange";

  my $new_mailbox = $Recipients->Create("organizationalPerson", "cn=$mailbox_alias_name");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating Mailbox -> $error_num ($error_name)\n",1);
    $Win32::OLE::Warn=$original_ole_warn_value;
    return 0;
  }
  my %attrs;
  $attrs{'uid'}=$mailbox_alias_name;
  $attrs{'mailPreferenceOption'}="0";
  $attrs{'MAPI-Recipient'}='TRUE'; 
  $attrs{'MDB-Use-Defaults'}="TRUE"; #By default set the box to adhere to Exchange Default settings
  $attrs{'givenName'}="Exchange"; #Temporary Name (it doesn't like returning from the subroutine without setting something)
  $attrs{'sn'}="Mailbox"; #Temporary Name
  $attrs{'cn'}="Exchange $mailbox_alias_name Mailbox";#Temporary Name
  $attrs{'Home-MTA'}="cn=Microsoft MTA,cn=$information_store_server,cn=Servers,cn=Configuration,ou=$ou,o=$org";
  $attrs{'Home-MDB'}="cn=Microsoft Private MDB,cn=$information_store_server,cn=Servers,cn=Configuration,ou=$ou,o=$org"; 

  foreach my $attr (keys %attrs) {
    $new_mailbox->Put($attr => $attrs{$attr}); 
  }
  $new_mailbox->SetInfo;
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting attribute on mailbox -> $error_num ($error_name)\n",1);
    $Win32::OLE::Warn=$original_ole_warn_value;
    return 0;
  }
  
  _DebugComment("      -Mailbox created...\n",3);

  $Win32::OLE::Warn=$original_ole_warn_value;
  return $new_mailbox;
}

sub _E2KCreateMailbox {
  my $error_num;
  my $error_name;
  my $provider;
  my $info_store_server;
  my $pdc;
  my $nt_pdc;
  my $mailbox_alias_name;
  my $mail_domain;
  my $storage_group;
  my $mb_store;
  my $mailbox_ldap_path;
  if (scalar(@_) >4) {
    $provider = $_[0];
    $info_store_server = $_[1];
    $nt_pdc = $_[2];
    $mailbox_alias_name = $_[3];
    if (scalar(@_) == 5) {
      #placeholder..
    } elsif (scalar(@_) == 6) {
      $mailbox_ldap_path = $_[5]
    } elsif (scalar(@_) == 7) {
      $storage_group = $_[5];
      $mb_store = $_[6];
    } else {
      _ReportArgError("CreateMailbox [E2K] (".scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("CreateMailbox [E2K] (".scalar(@_));
    return 0;
  }
  _StripBackslashes($nt_pdc,$pdc); 
  my $user_dist_name;
  if (!AdodbSearch($pdc,"(samAccountName=$mailbox_alias_name)","samAccountName,distinguishedName",$user_dist_name)) {
    _DebugComment("Error querying distinguished name for user in CreateMailbox (E2K)\n",1);
    return 0;
  }
 
  _DebugComment("user_dist_name = $user_dist_name\n",3);  
 
  bless $provider,"Win32::OLE";
  my $user_account = $provider->DataSource->Open("LDAP://$pdc/$user_dist_name",undef,adModeReadWrite);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening NT user account for new mailbox creation on $pdc ($error_num)\n",1);
    return 0;
  }
  my $info_store = $provider->GetInterface( "IMailboxStore");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening mailbox interface on $pdc ($error_num)\n",1);
    if ($error_num eq "0x80004002") {

      _DebugComment("Error:  No such interface supported.\n  Note:  Make sure you have the Exchange System Manager loaded on this system\n",2);
    }
    return 0;
  }
  if ($mailbox_ldap_path eq "") {
    if (!LocateMailboxStore($info_store_server,$storage_group,$mb_store,$mailbox_ldap_path)) {
      return 0;
    }
  }
  _DebugComment("$mailbox_ldap_path\n",3);
  $info_store->CreateMailbox($mailbox_ldap_path);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed creating mailbox for $mailbox_alias_name ($error_num) $error_name\n",1);
    return 0;
  }
 
  # SP2 Fix for perms issue could eventually be a problem:
  #   oObject.DataSource.Open strSourceURL, , adModeReadWrite
  #   http://support.microsoft.com/default.aspx?scid=kb;EN-US;q321039
  #the current implementation doesn't seem to have a problem with this
  #i.e. $provider->DataSource->Save(); #may eventually yield an error
  
  $provider->DataSource->Save();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed saving mailbox for $mailbox_alias_name ($error_num) $error_name\n",1);
    return 0;
  }
  return $provider;
}

sub GetDistinguishedName {
  my $server_name;
  my $filter;
  my $filter_name;
  my $result;
  if (scalar(@_) == 3) {
    $server_name = $_[0];
    $filter = $_[1];  
  } else {
    _ReportArgError("GetDistinguishedName",scalar(@_));
  }
  my %filters;
  
  %filters = ('Home-MDB' => "(objectClass=MHS-Message-Store)",
              'Home-MTA' => "(objectClass=MTA)",
           );
  if ($filters{$filter} ne "") {
    $filter_name=$filters{$filter};
  } else {
    $filter_name = $filter;#If someone wants to actually send a correctly formatted objectClass  
  }
  _DebugComment("filter=$filter_name\n",2);
  _DebugComment("search=$server_name\n",2);
  if (_AdodbExtendedSearch($server_name,"LDAP://$server_name",$filter_name,"cn,distinguishedName",2,$result)) {
    $_[2] = $result;
    return 1;
  } else {
    return 0;
  }
}

sub _StripBackslashes {
  my $nt_pdc = $_[0];
  if ($nt_pdc =~ /^\\\\(.*)/) {
    $_[1] = $1;
    return 1;
  } else {
    $_[1] = $nt_pdc;
    return 1;
  }
}

sub GetMailbox {
  my $error_num;
  my $error_name;
  my $mbx;
  my $provider = $_[0];

  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object for Exchange Server Determination for CreateMailbox\n",1);
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
  if ($type eq "IPerson") {
    #IPerson returns for CDO.Person (E2K)
    if ($mbx = _E2KGetMailbox(@_)) {
      bless $mbx,"Win32::Exchange";
      return $mbx;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($mbx = _E55GetMailbox(@_)) {
      bless $mbx,"Win32::Exchange";
      return $mbx;
    }
  }
  bless $provider,"Win32::Exchange";
  return 0;
}

sub _E55GetMailbox {
  my $ldap_provider;
  my $information_store_server;
  my $mailbox_alias_name;
  my $org;
  my $ou;
  my $error_num;
  my $error_name;
  my $find_mb;
  if (scalar(@_) > 2) {
    $ldap_provider = $_[0];
    $information_store_server = $_[1];
    $mailbox_alias_name = $_[2];
    if (scalar(@_) == 3) {
      if ($ldap_provider->GetLDAPPath($information_store_server,$org,$ou)) {
        _DebugComment("returned -> o=$org,ou=$ou\n",3);
      } else {
        _DebugComment("Error Returning from GetLDAPPath\n",1);
        return 0;
      }
    } elsif (scalar(@_) == 4) {
      $find_mb = $_[3];
    } elsif (scalar(@_) == 5) {
      $org = $_[3];
      $ou = $_[4];
    } else {
      _ReportArgError("GetMailbox [5.5]",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("GetMailbox [5.5] ",scalar(@_));
    return 0;
  }
  my $recipients_path;
  my $exch_mb_dn;
  if ($find_mb == 1) {
    if (_AdodbExtendedSearch($mailbox_alias_name,"LDAP://$information_store_server","(objectClass=organizationalPerson)","cn,distinguishedName",1,$exch_mb_dn)) {
      $recipients_path = "LDAP://$information_store_server/$exch_mb_dn";
    } else {
      _DebugComment("Error locating Exchange DL on the server.  Member addition cannot proceed.\n",1);
      return 0;
    }
  } else {
    $recipients_path = "LDAP://$information_store_server/cn=$mailbox_alias_name,cn=Recipients,ou=$ou,o=$org";
  }
  
  bless $ldap_provider,"Win32::OLE";
  my $Recipients = $ldap_provider->GetObject("",$recipients_path);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening recipients path on $information_store_server\n",1);
    return 0;
  }

  my $original_ole_warn_value = $Win32::OLE::Warn;
  $Win32::OLE::Warn = 0; #Turn STDERR warnings off because we probably are going to get an error (0x80072030)

  my $mailbox = $Recipients->GetObject("organizationalPerson", "cn=$mailbox_alias_name");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Unable to Get the mailbox object for $mailbox_alias_name on $information_store_server ($error_num)\n",1);
    $Win32::OLE::Warn=$original_ole_warn_value;
    return 0;
  }
  $Win32::OLE::Warn=$original_ole_warn_value;
  return $mailbox;
}

sub _E2KGetMailbox {
  my $error_num;
  my $error_name;
  my $provider;
  my $mailbox_alias_name;
  my $nt_dc;
  my $dc;
  if (scalar(@_) == 3) {
    $provider = $_[0];
    $nt_dc = $_[1];
    $mailbox_alias_name = $_[2];
  } else {
    _ReportArgError("GetMailbox [E2K]",scalar(@_));
    return 0;
  }
  _StripBackslashes ($nt_dc,$dc);
  
  my $user_dist_name;
  if (!AdodbSearch($dc,"(samAccountName=$mailbox_alias_name)","samAccountName,distinguishedName",$user_dist_name)) {
    _DebugComment("Error querying distinguished name for user in GetMailbox (E2K)\n",1);
    return 0;
  }

  bless $provider,"Win32::OLE";
  $provider->DataSource->Open("LDAP://$dc/$user_dist_name");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening AD user account for mailbox retrieval on $dc ($error_num)\n",1);
    return 0;
  }
  my $user_obj_path = $provider->DataSource->{SourceURL};
  my $user_obj = Win32::OLE->GetObject($user_obj_path);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening SourceURL for GetMailbox ($error_num)\n",1);
    return 0;
  }
  if ($user_obj->{homeMDB} eq "") {
    #Win32::OLE->LastError("0x80072030"); #This didn't work
    _DebugComment("Error performing GetMailbox..  mailbox does not exist ($error_num)\n",2);
    return 0;
  } else {
    $provider->DataSource->Save();
    return $provider;
  }
}

sub SetAttributes {
  my $error_num;
  my $error_name;
  my $provider = $_[0];

  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object type for Exchange Server Determination during call to SetAttributes\n",1);
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KSetAttributes(@_)) {
      bless $provider,"Win32::Exchange";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55SetAttributes(@_)) {
      bless $provider,"Win32::Exchange";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange";
  return 0;
}

sub _E55SetAttributes {
  my $error_num;
  my $error_name;
  my $mailbox;
  my %attrs;
  if (scalar(@_) == 2) {
    $mailbox = $_[0];
    if (ref($_[1]) ne "HASH") {
      _DebugComment("second object passed to SetAttributes was not a HASH reference -> $error_num ($error_name)\n",1);
      return 0;
    } else {
      %attrs = %{$_[1]};
    }
  } else {
    _ReportArgError("SetAttributes [E55]",scalar(@_));
    return 0;
  }
  my $original_ole_warn_value=$Win32::OLE::Warn;
  $Win32::OLE::Warn=0;
  bless $mailbox,"Win32::OLE";
  foreach my $attr (keys %attrs) {
    $mailbox->Put($attr => $attrs{$attr}); 
  }
  $mailbox->SetInfo(); 
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting attribute on mailbox -> $error_num ($error_name)\n",1);
    $Win32::OLE::Warn=$original_ole_warn_value;
    return 0;
  }
  $Win32::OLE::Warn=$original_ole_warn_value;
  return 1;
}

sub _E2KSetAttributes {
  my $error_num;
  my $error_name;
  my %attrs;
  my $user_account;
  my $mailbox;
  if (scalar(@_) == 2) {
    $user_account = $_[0];
    if (ref($_[1]) ne "HASH") {
      _DebugComment("second object passed to SetAttributes was not a HASH reference -> $error_num ($error_name)\n",1);
      return 0;
    } else {
      %attrs = %{$_[1]};
    }
  } else {
    _ReportArgError("SetAttributes [2K]",scalar(@_));
    return 0;
  }
  bless $user_account,"Win32::OLE";
  foreach my $interface (keys %attrs) {
    my $mailbox_interface = $user_account->GetInterface($interface);
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("error getting mailbox interface -> $error_num ($error_name)\n",1);
      return 0;
    }
    foreach my $attr (keys %{$attrs{$interface}}) {
      $mailbox_interface->{$attr} = $attrs{$interface}{$attr}; 
    }
    $user_account->DataSource->Save();
  }
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting attribute on mailbox -> $error_num ($error_name)\n",1);
    return 0;
  }
  return 1;

  #  overriding defaults
  #http://www.microsoft.com/technet/treeview/default.asp?url=/technet/prodtechnol/exchange/exchange2000/maintain/featusability/EX2KWSH.asp
  #  storage limits
  #http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wss/wss/_cdo_imailboxstore_interface.asp
  #  proxy addresses
  #http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wss/wss/_cdo_setting_proxy_addresses.asp
  #  interfaces and attributes:
  #http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wss/wss/_cdo_recipient_management_interfaces.asp
}

sub GetOwner {
  my $error_num;
  my $error_name;
  my $provider = $_[0];

  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object type for GetOwner\n",1);
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
  
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    #no available support for this operation
    bless $provider,"Win32::Exchange";
    return 0;
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55GetOwner(@_)) {
      bless $provider,"Win32::Exchange";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange";
  return 0;
}

sub _E55GetOwner {
  my $error_num;
  my $error_name;
  my $mailbox;
  my $returned_sid_type;
  if (scalar(@_) > 1) {
    $mailbox = $_[0];
    if (scalar(@_) == 2) {
      $returned_sid_type = ADS_SID_WINNT_PATH;    
    } elsif (scalar(@_) == 3) {
      $returned_sid_type = $_[2];    
    } else {
      _ReportArgError("GetOwner [5.5]",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("GetOwner [5.5]",scalar(@_));
    return 0;
  }


  bless $mailbox,"Win32::OLE";

  my $sid = Win32::OLE->new("ADsSID");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating ADsSID object -> $error_num ($error_name)\n",1);
    return 0;
  }
  $mailbox->GetInfoEx(["Assoc-NT-Account"],0);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error populating the property cache for Assoc-NT-Account -> $error_num ($error_name)\n",1);
    return 0;
  }

  $sid->SetAs(ADS_SID_HEXSTRING,$mailbox->{'Assoc-NT-Account'});
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating ADsSID object -> $error_num ($error_name)\n",1);
    return 0;
  }

  my $siduser = $sid->GetAs($returned_sid_type);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    if (ErrorCheck("0x80070534",$error_num,$error_name)) {
      _DebugComment("there was an error validating the SID from the Domain Controller (the account doesn't seem to exist anymore) -> $error_num ($error_name)\n",1);
      return 0;
    }
    _DebugComment("error getting SID to prepare for output -> $error_num ($error_name)\n",1);
    return 0;
  }
  $_[1] = $siduser;
  return 1;
}

sub SetOwner {
  my $error_num;
  my $error_name;
  my $dc;
  if (scalar(@_) != 2) {
    _ReportArgError("SetOwner [5.5]",scalar(@_));
    return 0;
  }
  my $new_mailbox = $_[0];
  my $username = $_[1];

  if ($username  =~ /(.*)\\(.*)/) {
    #DOMAIN\Username
    $dc=$1;
    $username = $2;
  } else {
    _DebugComment("error parsing username to extract domain and username\n",1);
    return 0;
  }

  my $sid = Win32::OLE->new("ADsSID");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating security object (ADsSID) -> $error_num ($error_name)\n",1);
    return 0;
  }
  $sid->SetAs(ADS_SID_WINNT_PATH, "WinNT://$dc/$username,user");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting security object at an ADS_SID_WINNT_PATH -> $error_num ($error_name)\n",1);
    return 0;
  }

  my $sidHex = $sid->GetAs(ADS_SID_HEXSTRING);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error converting security object at an ADS_SID_HEXSTRING -> $error_num ($error_name)\n",1);
    return 0;
  }

  bless $new_mailbox,"Win32::OLE";
  $new_mailbox->Put("Assoc-NT-Account", $sidHex );
  $new_mailbox->SetInfo;

  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting owner information on mailbox -> $error_num ($error_name)\n",1);
    bless $new_mailbox,"Win32::Exchange";
    return 0;      
  }
  bless $new_mailbox,"Win32::Exchange";
  return 1;
}

sub _E55GetPerms {
  #Need to work on this.
  if (scalar(@_) != 2) {
    _ReportArgError("GetPerms [5.5]",scalar(@_));
    return 0;
  }
  if (ref($_[1]) ne "ARRAY") {
    _DebugComment("permissions list must be an array reference (e55)\n",1);
    return 0;
  }
  my $mailbox = $_[0];

  my $sec = Win32::OLE->CreateObject("ADsSecurity");
  my $error_num;
  my $error_name;
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating security object (ADSSecurity) -> $error_num ($error_name)\n",1);
    if ($error_num eq "0x80004002") {
      _DebugComment("Error:  No such interface supported.\n  Note:  Make sure you have the ADSSecurity.DLL from the ADSI SDK regisered on this system\n",2);
    }
    return 0;
  }

  bless $mailbox,"Win32::OLE";
  
  my $sd = $sec->GetSecurityDescriptor($mailbox->{ADsPath});
  
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying security descriptor for mailbox -> $error_num ($error_name)\n",1);
    return 0;
  }
  my $dacl = $sd->{DiscretionaryAcl};
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying discretionary acl for mailbox -> $error_num ($error_name)\n",1);
    return 0;
  }
  @{$_[1]} = Win32::OLE::in($dacl);
  return 1;
}

sub SetPerms {
  my $error_num;
  my $error_name;
  my $provider = $_[0];

  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object type for Exchange Server Determination during call to SetAttributes\n",1);
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
  
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KSetPerms(@_)) {
      bless $provider,"Win32::Exchange";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55SetPerms(@_)) {
      bless $provider,"Win32::Exchange";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange";
  return 0;
}

sub _E55SetPerms {
  if (scalar(@_) != 2) {
    _ReportArgError("SetPerms [5.5]",scalar(@_));
    return 0;
  }
  if (ref($_[1]) ne "ARRAY") {
    _DebugComment("permissions list must be an array reference (e55)\n",1);
    return 0;
  }
  my $new_mailbox = $_[0];
  my @perms_list = @{$_[1]};

  my $sec = Win32::OLE->CreateObject("ADsSecurity");
  my $error_num;
  my $error_name;
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating security object (ADSSecurity) -> $error_num ($error_name)\n",1);
    if ($error_num eq "0x80004002") {
      _DebugComment("Error:  No such interface supported.\n  Note:  Make sure you have the ADSSecurity.DLL from the ADSI SDK regisered on this system\n",2);
    }
    return 0;
  }

  bless $new_mailbox,"Win32::OLE";
  
  my $sd = $sec->GetSecurityDescriptor($new_mailbox->{ADsPath});
  
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying security descriptor for mailbox -> $error_num ($error_name)\n",1);
    return 0;
  }
  my $dacl = $sd->{DiscretionaryAcl};
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying discretionary acl for mailbox -> $error_num ($error_name)\n",1);
    return 0;
  }

  foreach my $userid (@perms_list) {
    _DebugComment("      -Setting perms for $userid\n",3);
    my $ace = Win32::OLE->CreateObject("AccessControlEntry");
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("error creating access control entry for mailbox -> $error_num ($error_name)\n",1);
      return 0;
    }

    my %properties;
    $properties{Trustee}=$userid;
    $properties{AccessMask}=ADS_RIGHT_EXCH_MODIFY_USER_ATT | ADS_RIGHT_EXCH_MAIL_SEND_AS | ADS_RIGHT_EXCH_MAIL_RECEIVE_AS;
    $properties{AceType}=ADS_ACETYPE_ACCESS_ALLOWED;

    foreach my $property (keys %properties) {
      $ace->LetProperty($property,$properties{$property}); 
      if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
        _DebugComment("error setting $property for mailbox -> $error_num ($error_name)\n",1);
        return 0;
      }
    }


    $dacl->AddAce($ace);
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("error adding access control entry to perms list -> $error_num ($error_name)\n",1);
      return 0;
    }
  }
  $sd->LetProperty("DiscretionaryAcl",$dacl); 
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting  discretionary acl on security security descriptor -> $error_num ($error_name)\n",1);
    return 0;
  }
  $sec->SetSecurityDescriptor($sd);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting security descriptor on security object -> $error_num ($error_name)\n",1);
    return 0;
  }
  $new_mailbox->SetInfo;
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting permissions on mailbox -> $error_num ($error_name)\n",1);
    return 0;
  }
  return 1;
}

sub _E2KSetPerms {
  my $error_num;
  my $error_name;
  if (scalar(@_) != 2) {
    _ReportArgError("SetPerms [2K]",scalar(@_));
    return 0;
  }
  if (ref($_[1]) ne "ARRAY") {
    _DebugComment("permissions list must be an array reference (e2k)\n",1);
    return 0;
  }

  my $cdo_user_obj = $_[0];
  my @perms_list = @{$_[1]};

  bless $cdo_user_obj,"Win32::OLE";

  my $ldap_user_path = $cdo_user_obj->{DataSource}->{SourceURL};
  my $ldap_user_obj = Win32::OLE->GetObject($ldap_user_path);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error querying Source URL for CDO.Person object ($error_num)\n",1);
    return 0;
  }

  #http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q310866
  my $sd = $ldap_user_obj->{'MailboxRights'};
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error querying MailboxRights property ($error_num)\n",1);
    _DebugComment("- make sure you are using Exchange 2000 SP1+hotfix or higher [server & client]\n",2);
    _DebugComment('  http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q302926'."\n",3);
    return 0;
  }

  my $dacl = $sd->{DiscretionaryAcl};
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error getting DiscretionaryAcl ($error_num)\n",1);
    return 0;
  }

  foreach my $user_account (@perms_list) {
    my $domain;
    my $username;

    if ($user_account =~ /(.*)\\(.*)/) {
      $domain = $1;
      $username = $2;
    } else {
      _DebugComment("error parsing user object (expected DOMAIN\\Username) -> $error_num ($error_name)\n",1);
      return 0;
    }
    my $Ace = Win32::OLE->new("AccessControlEntry");
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Error creating new ACE ($error_num)\n",1);
      return 0;
    }
    my %properties;
    $properties{AccessMask}=ADS_RIGHT_DS_CREATE_CHILD;
    $properties{AceType}=ADS_ACETYPE_ACCESS_ALLOWED;
    $properties{AceFlags}=ADS_ACEFLAG_INHERIT_ACE;
    $properties{Flags}=0;
    $properties{Trustee}=$user_account;
    $properties{ObjectType}=0;
    $properties{InheritedObjectType}=0;
    foreach my $property (keys %properties) {
      if ($property =~ /(ObjectType|InheritedObjectType)/ && $properties{$property} == 0) {
        next;
      }
  
      $Ace->LetProperty($property,$properties{$property});
      if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
        _DebugComment("Error setting $property ($error_num)\n",1);
        return 0;
      }
    }
    $dacl->AddAce($Ace);
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Error adding AccessControlEntry to AccessControlList: ($error_num)\n",1);
      return 0;
    }
  }
  $sd->LetProperty('DiscretionaryAcl',$dacl);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error setting AccessControlList to Security Descriptor: ($error_num)\n",1);
    return 0;
  }
  $ldap_user_obj->LetProperty('MailboxRights',$sd);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error modfying Mailbox Security entry: ($error_num)\n",1);
    return 0;
  }
  $ldap_user_obj->SetInfo();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error setting information to Mailbox Security entry: ($error_num)\n",1);
    return 0;
  }
  return 1;
}

sub AddDLMembers {
  my $error_num;
  my $error_name;
  my $provider = $_[0];
  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object type for Exchange Server Determination during call to SetAttributes\n",1);
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
  
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KAddDLMembers(@_)) {
      bless $provider,"Win32::Exchange";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55AddDLMembers(@_)) {
      bless $provider,"Win32::Exchange";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange";
  return 0;
}

sub _E55AddDLMembers {
  my $ldap_provider;
  my $server_name;
  my $exch_dl_name;
  my @new_members;
  my $ou;
  my $org;
  my $find_dl;
  if (scalar(@_) > 3) {
    $ldap_provider = $_[0];
    $server_name=$_[1];
    $exch_dl_name=$_[2];
    if (ref($_[3]) ne "ARRAY") {
      _DebugComment("members list must be an array reference\n",1);
      return 0;
    }
    @new_members=@{$_[3]};
    if (scalar(@_) < 6) {
      if ($ldap_provider->GetLDAPPath($server_name,$org,$ou)) {
        _DebugComment("returned -> o=$org,ou=$ou\n",3);
      } else {
        _DebugComment("Error Returning from GetLDAPPath\n",1);
        return 0;
      }
      if (scalar(@_) == 5) {
        $find_dl = $_[4];
      }
    } elsif (scalar(@_) == 6) {
      $org = $_[4];
      $ou = $_[5];
    } else {
      _ReportArgError("AddDLMembers [5.5]",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("AddDLMembers [5.5]",scalar(@_));
    return 0;
  }

  my $temp_exch_dl;
  my $original_ole_warn_value = $Win32::OLE::Warn;

  bless $ldap_provider,"Win32::OLE";
  my $exch_dl_dn;
  my $exch_dl_path;
  my $temp_dl_path;
  my $exch_dl;
  if ($exch_dl_name =~ /^cn=.*ou=.*o=.*/) {
    #a dn was sent
    $exch_dl_path = "LDAP://$server_name/$exch_dl_name";
    $exch_dl_dn = $exch_dl_name;
  } else {
    if ($find_dl == 1) {
      if (_AdodbExtendedSearch($exch_dl_name,"LDAP://$server_name","(objectClass=groupOfNames)","cn,distinguishedName",$exch_dl_dn)) {
        $exch_dl_path = "LDAP://$server_name/$exch_dl_dn";
      } else {
        _DebugComment("Error locating Exchange DL on the server.  Member addition cannot proceed.\n",1);
        return 0;
      }
    } else {
      #an alias was sent (name only, check default container)
      $exch_dl_path = "LDAP://$server_name/cn=$exch_dl_name,cn=Distribution Lists,ou=$ou,o=$org";
      $exch_dl_dn = "cn=$exch_dl_name,cn=Distribution Lists,ou=$ou,o=$org";
    }
  }
  $exch_dl = $ldap_provider->GetObject("",$exch_dl_path);

  my $error_num;
  my $error_name;
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying distribution list ($exch_dl_name) -> $error_num ($error_name)\n",1);
    return 0;
  }
  
  my $exch_members = $exch_dl->{'member'}; #get the list
  if (ref($exch_members) eq "ARRAY") {
    _DebugComment("      -Array (2 or more members exist)\n",3);
  } else {
    _DebugComment("      -(Less than 2 members are named in this distribution list)\n",3);
    my $temp_exch_dl=$exch_members;
    undef ($exch_members);
    if ($temp_exch_dl) {
      _DebugComment("      -1 member exists\n",3);
      #So push the existing name to the Array
      push (@$exch_members, $temp_exch_dl);
    } else {
      _DebugComment("      -0 members exists\n",3);
    }
  }
  my $exch_mb_dn;
  foreach my $username (@new_members) {
    _DebugComment("      -Adding $username to Distribution List: $exch_dl_name\n",2);
    if ($username =~ /^cn=.*ou=.*o=.*$/) {
      $exch_mb_dn = $username;
    } else {
      if ($find_dl == 1) {
        if (!_AdodbExtendedSearch($username,"LDAP://$server_name","(objectClass=organizationalPerson)","rdn,distinguishedName",$exch_mb_dn)) {
          _DebugComment("Error locating Exchange mailbox on the server.  Member addition cannot proceed.\n",1);
          return 0;
        }
      } else {
        $exch_mb_dn = "cn=$username,cn=Recipients,ou=$ou,o=$org";
      }
    }
    my $duplicate;
    foreach my $dup (@$exch_members) {
      if (lc($dup) eq lc($exch_mb_dn)) {
        _DebugComment("Error adding user ($username) to distribution list [they are already a member]\n",1);
        $duplicate = 1;
        last;
      }
    }
    if ($duplicate != 1) {
      push (@$exch_members, $exch_mb_dn);
    }
  }
  $exch_dl->Put('member', $exch_members);
  $exch_dl->SetInfo;
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting new member for distribution list ($exch_dl_name) -> $error_num ($error_name)\n",1);
    return 0;
  }
  return 1;
}

sub _E2KAddDLMembers {
  if (scalar(@_) != 3) {
    _ReportArgError("AddDLMembers (E2K)",scalar(@_));
    return 0;
  }
  if (ref($_[2]) ne "ARRAY") {
    _DebugComment("Third argument is the list of users you want to add to this DL, and should be an array reference, but instead, it was a(an): ".ref($_[2])." reference\n",1);
    return 0;
  }
  my $error_num;
  my $error_name;
  my $group_dn;
  my $user_dn;
  my $provider = $_[0];
  my $group = $_[1];
  my @user_list = @{$_[2]};

  if (!Win32::Exchange::_AdodbExtendedSearch($group,"LDAP://RootDSE/dnsHostName","(objectClass=group)","samAccountName,distinguishedName",$group_dn)) {
    _DebugComment("Failed Adodb search for dist list\n",1);
    return 0;
  }

  foreach my $username (@user_list) {
    #print "    -Adding $username to $group\n";
    if (!Win32::Exchange::_AdodbExtendedSearch($username,"LDAP://RootDSE/dnsHostName","(objectClass=user)","samAccountName,distinguishedName",$user_dn)) {
      _DebugComment("Failed Adodb search for user\n",1);
      return 0;
    }
    my $RootDSE = Win32::OLE->GetObject("LDAP://RootDSE");
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Failed Getting RootDSE ($error_num)\n",1);
      return 0;
    }
    my $dc = $RootDSE->Get("dnsHostName");
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Error getting RootDSE dns host name ($error_num)\n",1);
      return 0;
    }
    my @dc_array = split(/\./,$dc);
    my $ldap_obj = Win32::OLE->new("ADsNamespaces");
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Error creating new ADsNamespaces object ($error_num)\n",1);
      return 0;
    }
    my $group_obj = $ldap_obj->GetObject("","LDAP://$dc_array[0]/$group_dn");
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Error opening distribution list on $dc_array[0] ($error_num)\n",1);
      return 0;
    }
  
    $group_obj->Add("LDAP://$dc_array[0]/$user_dn");
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      if ($error_num eq "0x80071392") {
        _DebugComment("Error adding user ($username) to distribution list [they are already a member]\n",1);
      } else {
        _DebugComment("Error adding user ($username) to distribution list ($error_num)\n",1);
        return 0;
      }
    }
 
    $group_obj->SetInfo;
    if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
      _DebugComment("Error committing addition to distribution list ($error_num)\n",1);
      return 0;
    }
  }
  return 1;
}

sub ErrorCheck {
  my $last_error_expected = $_[0];
  my $error_num;
  my $error_name;
  my $last_ole_error = Win32::OLE->LastError();
  $error_num = sprintf ("0x%08x",$last_ole_error);
  my @error_list = split(/\"/,$last_ole_error,3);
  $error_name = $error_list[1];
  if ($error_num ne $last_error_expected) {
    $_[1] = $error_num;
    $_[2] = $error_name;
    return 0;
  } else {
    return 1;
  }
}

sub _ReportArgError {
  _DebugComment("incorrect number of options passed to $_[0] ($_[1])\n",0);
  return 1;
}

sub _DebugComment {
  if (scalar(@_) != 2) {
    print "DebugComment Error!!!!\n";
  }
  print "$_[0]" if ($DEBUG > ($_[1] - 1));
  return 1;
}

1;
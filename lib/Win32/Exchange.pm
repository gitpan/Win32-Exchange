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
#use vars qw ($VERSION $DEBUG);

use Win32::OLE;
Win32::OLE->Initialize(Win32::OLE::COINIT_OLEINITIALIZE);

#Win32::OLE->Option('_Unique' => 1);
#@ISA = qw(Win32::OLE);

my $VERSION = "0.0.0.021";
my $DEBUG = 1;

#CONSTANTS

###Various ADS Objects used in Mailbox Creation ##
my $ADS_SID_HEXSTRING = 0x01;                    #
my $ADS_SID_WINNT_PATH = 0x05;                   # 
my $ADS_RIGHT_EXCH_MODIFY_USER_ATT = 0x02;       # 
my $ADS_RIGHT_EXCH_MAIL_SEND_AS = 0x08;          #
my $ADS_RIGHT_EXCH_MAIL_RECEIVE_AS = 0x10;       #
my $ADS_ACETYPE_ACCESS_ALLOWED = 0x00;           #
##################################################

####Contsants used with OpenDSObject######  NOT USED YET
my $ADS_SECURE_AUTHENTICATION  = 0x1;    #
my $ADS_USE_ENCRYPTION         = 0x2;    #
my $ADS_USE_SSL                = 0x2;    #
my $ADS_READONLY_SERVER        = 0x4;    #
my $ADS_PROMPT_CREDENTIALS     = 0x8;    #
my $ADS_NO_AUTHENTICATION      = 0x10;   #
my $ADS_FAST_BIND              = 0x20;   #
my $ADS_USE_SIGNING            = 0x40;   #
my $ADS_USE_SEALING            = 0x80;   #
my $ADS_USE_DELEGATION         = 0x100;  #
my $ADS_SERVER_BIND            = 0x200;  #
##########################################



sub new {
  my $server;
  my $ver;
  my %version;
  if (scalar(@_) == 1) {
    $server = $_[0];
  } elsif (scalar(@_) == 2) {
    $server = $_[0];
    $ver = $_[1];
  } else {
    _ReportArgError("new",scalar(@_));
    return 0;
  }
  my $class = "Win32::Exchange";
  my $ldap_provider = {};
  if (scalar(@_) == 1) {
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
      _DebugComment("Failed creating ADsNamespaces object");
      return undef;
    }
  } elsif ($ver eq "6.0") {
    #Exchange 2000
    if ($ldap_provider = Win32::OLE->new('CDO.Person')) {
      return bless $ldap_provider,$class;
    } else {
      _DebugComment("Failed creating CDO.Person object\n");
      return undef;
    }
  } else {
    _DebugComment("ver not right\n");
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
    _DebugComment("result = $result\n");
    if ($result =~ /cn=.*,cn=Servers,cn=Configuration,ou=(.*),o=(.*)/) {
      my $returned_ou = $1;
      my $returned_o = $2;
      $_[$return_point]=$returned_o;
      $_[($return_point+1)]=$returned_ou;
      _DebugComment("ou=$returned_ou\no=$returned_o\n");
      return 1;
    } else {
      _DebugComment("result = $result\n");
      _DebugComment("result from ADODB search failed to produce an acceptable match\n");
      return 0;
    }
  } else {
    _DebugComment("ADODB search failed\n");
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
    $server_name = $_[1];
  } else {
    _ReportArgError("GetVersion",scalar(@_));
    return 0;
  }

  my $serial_val;
  my $serial_version_check_obj = Win32::OLE->new('CDOEXM.ExchangeServer'); #substantiates the possible existance of e2k
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    if ($error_num eq "0x80040154" ||
        $error_num eq "0x800401f3") {
      #0x80040154 Class not registered
      #0x800401f3 Invalid class string
      _DebugComment("The Exchange 2000 client tools don't look to be installed on this machine\n");
      if (!_E55VersionInfo($server_name,$serial_val)) {
        _DebugComment("Error getting version information from Exchange 5.5\n");
        return 0;
      }
    } else {
      _DebugComment("error: $error_num - $error_name on $server_name encountered while trying to perform GetVersion\n");
      return 0;
    }
  } else {
    _DebugComment("found e2k tools, so we'll look and see what version of Exchange you have.\n");
    if (!_E2kVersionInfo($server_name,$serial_val)) {
      _DebugComment("Error getting version information from Exchange 2000\n");
      return 0;
    }
  }

  if ($serial_val =~ /Version (.*) \(Build (.*): Service Pack (.*)\)/i) {
    my %return_struct;
    $return_struct{ver}= $1;
    $return_struct{build}= $2;
    $return_struct{sp}= $3;
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
    $server_name = $_[0];
  } else {
    _ReportArgError("_E55VersionInfo",scalar(@_));
    return 0;
  }
  my $serial_val;
  if (_AdodbExtendedSearch($server_name,"rootdse-configurationnamingcontext","(objectCategory=msExchExchangeServer)","name,serialNumber",$serial_val)) {
    if ($serial_val =~ /Version (.*) \(Build (.*): Service Pack (.*)\)/i) {
      $_[1] = $serial_val;
      return 1;
    } else {
      _DebugComment("GetVersion failed to produce acceptable results (E55)\n");
      return 0;
    }
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
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("Failed creating object for version information (E2K) on $server_name -> $error_num ($error_name)\n");
      return 0;
  }
  $exchange_server->DataSource->Open($server_name);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("Failed opening object for version information (E2K) on $server_name -> $error_num ($error_name)\n");
      return 0;
  }

  #exampe output:
  #Version 5.5 (Build 2653.23: Service Pack 4)
  #Version 6.0 (Build 6249.4: Service Pack 3)

  if ($exchange_server->{ExchangeVersion} ne "") {
    $_[1] = $exchange_server->{ExchangeVersion};
    return 1;
  } else {
    _DebugComment("Failed failed to produce valid version info for $server_name\n");
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
      _ReportArgError("_AdodbSearch",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("_AdodbSearch",scalar(@_));
    return 0;
  }
  my @ado_columns = split (/,/,$columns);
  my $Conn = Win32::OLE->new("ADODB.Connection");
  $Conn->{'Provider'} = "ADsDSOObject";
  $Conn->Win32::OLE::Open("Active Directory Provider;UID=;PWD=");
  my $path = "<$ldap_path>;$filter;$columns;subtree";
  my $RS = $Conn->Win32::OLE::Execute($path);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("path=$ldap_path\nfilter=$filter\ncolumns=$columns\nFailed Executing ADODB Execute command on $server_name -> $error_num ($error_name)\n");
    return 0;
  }
  if ($RS->RecordCount == 0) {
    _DebugComment("path=$ldap_path\nfilter=$filter\ncolumns=$columns\nAdodbSearch yeilded no results for search on $server_name -> $error_num ($error_name)\n");
  } elsif ($RS->RecordCount > 1) {
    _DebugComment("path=$ldap_path\nfilter=$filter\ncolumns=$columns\nAdodbSearch yeilded more than 1 result for search on $server_name -> $error_num ($error_name)\n");
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
    _DebugComment("Only 2 columns can be sent to _AdodbExtendedSearch (total recieved = ".scalar(@cols).")\n");
  }
  if (lc($path) eq "rootdse-configurationnamingcontext") {
    my $RootDSE = Win32::OLE->GetObject("LDAP://RootDSE");
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
        _DebugComment("Failed creating object for _AdodbExtendedSearch on $server_name -> $error_num ($error_name)\n");
        return 0;
    }
    $path = "LDAP://".$RootDSE->Get("configurationNamingContext");
  }
  my $string = "<$path>;$filter;$columns;subtree";
  my $Com = Win32::OLE->new("ADODB.Command");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("path=$path\nfilter=$filter\ncolumns=$columns\nFailed creating ADODB.Command object for _AdodbExtendedSearch on $server_name -> $error_num ($error_name)\n");
      return 0;
  }
  my $Conn = Win32::OLE->new("ADODB.Connection");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("path=$path\nfilter=$filter\ncolumns=$columns\nFailed creating ADODB.Connection object for version information (E55) on $server_name -> $error_num ($error_name)\n");
      return 0;
  }
  $Conn->{'Provider'} = "ADsDSOObject";
  $Conn->Open("ADs Provider");
  $Com->{ActiveConnection} = $Conn;
  my $RS = $Conn->Execute($string);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("path=$path\nfilter=$filter\ncolumns=$columns\nFailed executing ADODB.Command for version information (E55) on $server_name -> $error_num ($error_name)\n");
      return 0;
  }
  my $not_found = 1;
  my $search_val;
  while ($not_found == 1) {
    if ($fuzzy == 1) {
      if ($RS->Fields($cols[1])->value =~ /$server_name/i) {
        if (ref($RS->Fields($cols[1])->value) eq "ARRAY") {
          $search_val = @{$RS->Fields($cols[1])->value}[0]; 
        } else {
          $search_val = $RS->Fields($cols[1])->value; 
        }
        $not_found = 0;
      }
    } else {
      if ($server_name eq $RS->Fields($cols[0])->value) {
        if (ref($RS->Fields($cols[1])->value) eq "ARRAY") {
          $search_val = @{$RS->Fields($cols[1])->value}[0]; 
        } else {
          $search_val = $RS->Fields($cols[1])->value; 
        }
        $not_found = 0;
      }
    }
    if ($RS->EOF) {
      $not_found = -1;
    }
    if ($not_found == 1) {
      print $RS->Fields($cols[1])->value."\n";
      $RS->MoveNext;
    }        
  }
  if ($not_found == -1) {
    _DebugComment("Unable to match valid data for your search on $server_name\n");
    return 0;
  }
  $_[$return_point] = $search_val;
  return 1;
}

sub LocateMailboxStore {
  my $store_server;
  my $storage_group;
  my $mb_store;
  my $count = "no";
  if (scalar(@_) > 3) {
    if (scalar(@_) == 4) {
    } elsif (scalar(@_) == 5) {
      if (ref($_[4]) eq "ARRAY") {
        $count = "yes";
      } else {
        _DebugComment("the fifth argument passed to LocateMailboxStore must be an array (but is optional).\n");
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
      _DebugComment("Unable to locate valid mailbox store for mailbox creation.\n");
      return 0;          
    }
  } else {
    _DebugComment("Unable to locate valid storage group for mailbox creation.\n");
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
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening ADODB ExchangeServer object for Storage Group enumeration on $server_name -> $error_num ($error_name)\n");
    return 0;
  }

  my @storegroups = Win32::OLE::in($exchange_server->StorageGroups);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed enumerating Storage Groups on $server_name -> $error_num ($error_name)\n");
    return 0;
  }
  my %storage_groups;
  my $stor_group_obj = Win32::OLE->new("CDOEXM.StorageGroup");
  my $mbx_store_obj = Win32::OLE->new("CDOEXM.MailboxStoreDB");
  foreach my $storegroup (@storegroups) {
    $stor_group_obj->DataSource->Open($storegroup);
    _DebugComment("Stor Name = ".$stor_group_obj->{Name}."\n");
    foreach my $mbx_store (Win32::OLE::in($stor_group_obj->{MailboxStoreDBs})) {
      $mbx_store_obj->DataSource->Open($mbx_store);
      _DebugComment("  Mailbox Store = $mbx_store_obj->{Name}\n");
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
  if (ref(@_[0]) ne "HASH") {
    _DebugComment("Storage group object is not a hash\n");
    return 0;
  }
  my %storage_groups = %{$_[0]};
  my $info_store_server = $_[1];
  my $storage_group = $_[2];
  my $mb_store = $_[3];
  my $ldap_path;
  if (scalar(keys %storage_groups) == 0) {
      _DebugComment("No Storage Groups were found\n");
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
      _DebugComment("Error locating proper storage group and mailbox db for mailbox creation (1SG)\n");
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
      _DebugComment("Error locating proper storage group and mailbox db for mailbox creation (2+SG)\n");
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
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object for Exchange Server Determination for CreateMailbox ($error_num)\n");
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
  if ($type eq "IPerson") {
    #IPerson returns for CDO.Person (E2K)
    if ($mbx = _E2KCreateMailbox(@_)) {
      return $mbx;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($mbx = _E55CreateMailbox(@_)) {
      return $mbx;
    }
  }
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
  if (scalar(@_) > 2) {
    $ldap_provider = $_[0];
    $information_store_server = $_[1];
    $mailbox_alias_name = $_[2];
    if (scalar(@_) == 3) {
      if ($ldap_provider->GetLDAPPath($information_store_server,$org,$ou)) {
        _DebugComment("returned -> o=$org,ou=$ou\n");
      } else {
        _DebugComment("Error Returning from GetLDAPPath\n");
        return 0;
      }
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
  my $recipients_path = "LDAP://$information_store_server/cn=Recipients,ou=$ou,o=$org";
  _DebugComment("$recipients_path\n");
  bless $ldap_provider,"Win32::OLE";
  my $Recipients = $ldap_provider->GetObject("",$recipients_path);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening recipients path on $information_store_server\n");
    bless $ldap_provider,"Win32::Exchange";
    return 0;
  }

  my $original_ole_warn_value = $Win32::OLE::Warn;
  $Win32::OLE::Warn = 0; #Turn STDERR warnings off because we probably are going to get an error (0x80072030)

  $Recipients->GetObject("organizationalPerson", "cn=$mailbox_alias_name");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x80072030",$error_num,$error_name)) {
    if ($error_num eq "0x00000000") {
      _DebugComment("$error_num - Mailbox already exists on $information_store_server\n");
      $Win32::OLE::Warn=$original_ole_warn_value;
      bless $ldap_provider,"Win32::Exchange";
      return 0;
    } else {
      _DebugComment("Unable to lookup object $mailbox_alias_name on $information_store_server ($error_num)\n");
      $Win32::OLE::Warn=$original_ole_warn_value;
      bless $ldap_provider,"Win32::Exchange";
      return 0;
    }
  }
  _DebugComment("    Box Does Not Exist (This is good)\n");
  bless $ldap_provider,"Win32::Exchange";

  my $new_mailbox = $Recipients->Create("organizationalPerson", "cn=$mailbox_alias_name");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating Mailbox -> $error_num ($error_name)\n");
    $Win32::OLE::Warn=$original_ole_warn_value;
    bless $new_mailbox,"Win32::Exchange";
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
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting attribute on mailbox -> $error_num ($error_name)\n");
    $Win32::OLE::Warn=$original_ole_warn_value;
    bless $new_mailbox,"Win32::Exchange";
    return 0;
  }
  
  _DebugComment("      -Mailbox created...\n");

  $Win32::OLE::Warn=$original_ole_warn_value;
  return bless $new_mailbox,"Win32::Exchange";
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
  if (scalar(@_) >4) {
    $provider = $_[0];
    $info_store_server = $_[1];
    $nt_pdc = $_[2];
    $mailbox_alias_name = $_[3];
    if (scalar(@_) == 5) {
      #placeholder..
    } elsif (scalar(@_) == 7) {
      $storage_group = $_[4];
      $mb_store = $_[5];
    } else {
      _ReportArgError("CreateMailbox [E2K] (".scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("CreateMailbox [E2K] (".scalar(@_));
    return 0;
  }
  if ($nt_pdc =~ /^\\\\.*/) {
    $nt_pdc =~ /^\\\\(.*)/;
    $pdc = $1;
  } else {
    $pdc = $nt_pdc;
  }

  my $user_dist_name;
  if (!AdodbSearch($pdc,"(samAccountName=$mailbox_alias_name)","samAccountName,distinguishedName",$user_dist_name)) {
    _DebugComment("Error querying distinguished name for user in CreateMailbox (E2K)\n");
    return 0;
  }

  #_DebugComment("user_dist_name = $user_dist_name\n");  

  bless $provider,"Win32::OLE";
  my $user_account = $provider->DataSource->Open("LDAP://$pdc/$user_dist_name");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening NT user account for new mailbox creation on $pdc ($error_num)\n");
    bless $provider,"Win32::Exchange";
    return 0;
  }
  my $info_store = $provider->GetInterface( "IMailboxStore");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening mailbox interface on $pdc ($error_num)\n");
    if ($error_num eq "0x80004002") {
      _DebugComment("Error:  No such interface supported.\n  Note:  Make sure you have the Exchange System Manager loaded on this system\n");
    }
    bless $provider,"Win32::Exchange";
    return 0;
  }
  my $mailbox_ldap_path = "";
  if (!LocateMailboxStore($info_store_server,$storage_group,$mb_store,$mailbox_ldap_path)) {
    return 0;
  }
  _DebugComment("$mailbox_ldap_path\n");
  my $user_display_name;
  if (!AdodbSearch($pdc,"(samAccountName=$mailbox_alias_name)","samAccountName,displayName",$user_display_name)) {
    _DebugComment("Error querying distinguished name for user\n");
    return 0;
  }
  $info_store->CreateMailbox($mailbox_ldap_path);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed creating mailbox for $mailbox_alias_name ($error_num) $error_name\n");
    bless $provider,"Win32::Exchange";
    return 0;
  }

  # SP2 Fix for perms issue could eventually be a problem:
  #   oObject.DataSource.Open strSourceURL, , adModeReadWrite
  #   http://support.microsoft.com/default.aspx?scid=kb;EN-US;q321039
  #the current implementation doesn't seem to have a problem with this
  #i.e. $provider->DataSource->Save(); #may eventually yield an error
  
  $provider->DataSource->Save();
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed saving mailbox for $mailbox_alias_name\n");
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
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
  
  %filters ('Home-MDB' => "(objectClass=MHS-Message-Store)",
            'Home-MTA' => "(objectClass=MTA)",
           );
  if (defined($filters{$filter})) {
    $filter_name=$filters{$filter};
  } else {
    $filter_name = $filter;#If someone wants to actually send a correctly formatted objectClass  
  }
  if (_AdodbExtendedSearch($server_name,"LDAP://$server_name",$filter_name,"cn,distinguishedName",1,$result)) {
    $_[2] = $result;
    return 0;
  } else {
    return 0;
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
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object for Exchange Server Determination for CreateMailbox\n");
    bless $provider,"Win32::Exchange";
    return 0;
  }
  bless $provider,"Win32::Exchange";
  if ($type eq "IPerson") {
    _DebugComment("Not functional at this time.\nI don't know that this is a feasible sub since the object you want to look at is an User Account.\n");
    return 0;
    #IPerson returns for CDO.Person (E2K)
    if ($mbx = _E2KGetMailbox(@_)) {
      return $mbx;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($mbx = _E55GetMailbox(@_)) {
      return $mbx;
    }
  }
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
  if (scalar(@_) > 2) {
    $ldap_provider = $_[0];
    $information_store_server = $_[1];
    $mailbox_alias_name = $_[2];
    if (scalar(@_) == 3) {
      if ($ldap_provider->GetLDAPPath($information_store_server,$org,$ou)) {
        _DebugComment("returned -> o=$org,ou=$ou\n");
      } else {
        _DebugComment("Error Returning from GetLDAPPath\n");
        return 0;
      }
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
  my $recipients_path = "LDAP://$information_store_server/cn=Recipients,ou=$ou,o=$org";
  bless $ldap_provider,"Win32::OLE";
  my $Recipients = $ldap_provider->GetObject("",$recipients_path);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening recipients path on $information_store_server\n");
    bless $ldap_provider,"Win32::Exchange";
    return 0;
  }

  my $original_ole_warn_value = $Win32::OLE::Warn;
  $Win32::OLE::Warn = 0; #Turn STDERR warnings off because we probably are going to get an error (0x80072030)

  my $mailbox = $Recipients->GetObject("organizationalPerson", "cn=$mailbox_alias_name");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("Unable to Get the mailbox object for $mailbox_alias_name on $information_store_server ($error_num)\n");
    $Win32::OLE::Warn=$original_ole_warn_value;
    bless $ldap_provider,"Win32::Exchange";
    return 0;
  }
  $Win32::OLE::Warn=$original_ole_warn_value;
  return bless $mailbox,"Win32::Exchange";
}

#sub _E2KGetMailbox {
#  my $error_num;
#  my $error_name;
#  my $provider;
#  my $info_store_server;
#  my $pdc;
#  my $nt_pdc;
#  my $mailbox_alias_name;
#  my $mail_domain;
#  my $storage_group;
#  my $mb_store;
#  if (scalar(@_) >4) {
#    $provider = $_[0];
#    $info_store_server = $_[1];
#    $nt_pdc = $_[2];
#    $mailbox_alias_name = $_[3];
#    if (scalar(@_) == 5) {
#      #placeholder..
#    } elsif (scalar(@_) == 7) {
#      $storage_group = $_[4];
#      $mb_store = $_[5];
#    } else {
#      _ReportArgError("GetMailbox [E2K]",scalar(@_));
#      return 0;
#    }
#  } else {
#    _ReportArgError("GetMailbox [E2K]",scalar(@_));
#    return 0;
#  }
#  if ($nt_pdc =~ /^\\\\.*/) {
#    $nt_pdc =~ /^\\\\(.*)/;
#    $pdc = $1;
#  } else {
#    $pdc = $nt_pdc;
#  }
#
#  my $user_dist_name;
#  if (!AdodbSearch($pdc,"(samAccountName=$mailbox_alias_name)","samAccountName,distinguishedName",$user_dist_name)) {
#    _DebugComment("Error querying distinguished name for user in CreateMailbox (E2K)\n");
#    return 0;
#  }
#
#  bless $provider,"Win32::OLE";
#  my $user_account = $provider->DataSource->Open("LDAP://$pdc/$user_dist_name");
#  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
#    _DebugComment("Failed opening NT user account for new mailbox creation on $pdc ($error_num)\n");
#    bless $provider,"Win32::Exchange";
#    return 0;
#  }
#  my $info_store = $provider->GetInterface( "IMailboxStore");
#  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
#    _DebugComment("Failed opening mailbox interface on $pdc ($error_num)\n");
#    if ($error_num eq "0x80004002") {
#      _DebugComment("Error:  No such interface supported.\n  Note:  Make sure you have the Exchange System Manager loaded on this system\n");
#    }
#    bless $provider,"Win32::Exchange";
#    return 0;
#  }
#  #this isn't really a mailbox, but the underlying object that you need to do manipulations on.
#  #this isn't tested much though.  Sorry.
#  return $user_account;
#}

sub SetAttributes {
  my $error_num;
  my $error_name;
  my $mailbox = $_[0];

  bless $mailbox,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($mailbox);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object type for Exchange Server Determination during call to SetAttributes\n");
    bless $mailbox,"Win32::Exchange";
    return 0;
  }
  bless $mailbox,"Win32::Exchange";
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($mailbox = _SetAttributes_2K(@_)) {
      return $mailbox;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($mailbox = _SetAttributes_55(@_)) {
      return $mailbox;
    }
  }
  return 0;
}
sub _SetAttributes_55 {
  my $error_num;
  my $error_name;
  my $mailbox;
  my %attrs;
  if (scalar(@_) == 2) {
    $mailbox = $_[0];
    if (ref($_[1]) ne "HASH") {
      _DebugComment("second object passed to SetAttributes was not a HASH reference -> $error_num ($error_name)\n");
      return 0;
    }
    %attrs = %{$_[1]};
  } else {
    _ReportArgError("SetAttributes [5.5]",scalar(@_));
    return 0;
  }
  my $original_ole_warn_value=$Win32::OLE::Warn;
  $Win32::OLE::Warn=0;
  bless $mailbox,"Win32::OLE";
  foreach my $attr (keys %attrs) {
    $mailbox->Put($attr => $attrs{$attr}); 
  }
  $mailbox->SetInfo;
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting attribute on mailbox -> $error_num ($error_name)\n");
    $Win32::OLE::Warn=$original_ole_warn_value;
    bless $mailbox,"Win32::Exchange";
    return 0;
  }
  $Win32::OLE::Warn=$original_ole_warn_value;
  bless $mailbox,"Win32::Exchange";
  return 1;
}

sub _SetAttributes_2K {
  my $error_num;
  my $error_name;
  my %attrs;
  my $user_account;
  my $mailbox;
  if (scalar(@_) == 2) {
    $user_account = $_[0];
    if (ref($_[1]) ne "HASH") {
      _DebugComment("second object passed to SetAttributes was not a HASH reference -> $error_num ($error_name)\n");
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
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error getting mailbox interface -> $error_num ($error_name)\n");
      bless $user_account,"Win32::Exchange";
      bless $mailbox_interface,"Win32::Exchange";
      return 0;
    }
    foreach my $attr (keys %{$attrs{$interface}}) {
      $mailbox_interface->{$attr} = $attrs{$interface}{$attr}; 
    }
    $user_account->DataSource->Save();
  }
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting attribute on mailbox -> $error_num ($error_name)\n");
    bless $mailbox,"Win32::Exchange";
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
    _DebugComment("error parsing username to extract domain and username\n");
    return 0;
  }

  my $sid = Win32::OLE->new("ADsSID");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating security object (ADsSID) -> $error_num ($error_name)\n");
    return 0;
  }
  $sid->SetAs($ADS_SID_WINNT_PATH, "WinNT://$dc/$username,user");
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting security object at an ADS_SID_WINNT_PATH -> $error_num ($error_name)\n");
    return 0;
  }

  my $sidHex = $sid->GetAs($ADS_SID_HEXSTRING);
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error converting security object at an ADS_SID_HEXSTRING -> $error_num ($error_name)\n");
    return 0;
  }

  bless $new_mailbox,"Win32::OLE";
  $new_mailbox->Put("Assoc-NT-Account", $sidHex );
  $new_mailbox->SetInfo;

  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting owner information on mailbox -> $error_num ($error_name)\n");
    bless $new_mailbox,"Win32::Exchange";
    return 0;      
  }
  bless $new_mailbox,"Win32::Exchange";
  return 1;
}

sub SetPerms {
  if (scalar(@_) != 2) {
    _ReportArgError("SetPerms [5.5]",scalar(@_));
    return 0;
  }
  if (ref($_[1]) ne "ARRAY") {
    _DebugComment("permissions list must be an array reference\n");
    return 0;
  }
  my $new_mailbox = $_[0];
  my @perms_list = @{$_[1]};

  my $sec = Win32::OLE->CreateObject("ADsSecurity");
  my $error_num;
  my $error_name;
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating security object (ADSSecurity) -> $error_num ($error_name)\n");
    if ($error_num eq "0x80004002") {
      _DebugComment("Error:  No such interface supported.\n  Note:  Make sure you have the ADSSecurity.DLL from the ADSI SDK regisered on this system\n");
    }
    return 0;
  }

  bless $new_mailbox,"Win32::OLE";
  
  my $sd = $sec->GetSecurityDescriptor($new_mailbox->{ADsPath});
  
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying security descriptor for mailbox -> $error_num ($error_name)\n");
    bless $new_mailbox,"Win32::Exchange";
    return 0;
  }
  my $dacl = $sd->{DiscretionaryAcl};
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying discretionary acl for mailbox -> $error_num ($error_name)\n");
    bless $new_mailbox,"Win32::Exchange";
    return 0;
  }

  foreach my $userid (@perms_list) {
    _DebugComment("      -Setting perms for $userid\n");
    my $ace = Win32::OLE->CreateObject("AccessControlEntry");
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error creating access control entry for mailbox -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }

    $ace->LetProperty('Trustee',$userid); 
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting trustee for mailbox -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }

    $ace->LetProperty('AccessMask',$ADS_RIGHT_EXCH_MODIFY_USER_ATT | $ADS_RIGHT_EXCH_MAIL_SEND_AS | $ADS_RIGHT_EXCH_MAIL_RECEIVE_AS);
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting access mask for mailbox -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }

    $ace->LetProperty('AceType', $ADS_ACETYPE_ACCESS_ALLOWED);
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting ace type for mailbox -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }
    $dacl->AddAce($ace);
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error adding access control entry to perms list -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }

    $sd->LetProperty("DiscretionaryAcl",$dacl); 
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting  discretionary acl on security security descriptor -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }
    $sec->SetSecurityDescriptor($sd);
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting security descriptor on security object -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }
  }
  $new_mailbox->SetInfo;
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting permissions on mailbox -> $error_num ($error_name)\n");
    bless $new_mailbox,"Win32::Exchange";
    return 0;
  }
  bless $new_mailbox,"Win32::Exchange";
  return 1;
}

sub _E2KSetPerms {
  if (scalar(@_) != 2) {
    _ReportArgError("SetPerms [2K]",scalar(@_));
    return 0;
  }
  if (ref($_[1]) ne "ARRAY") {
    _DebugComment("permissions list must be an array reference\n");
    return 0;
  }
  my $new_mailbox = $_[0];
  my @perms_list = @{$_[1]};

  my $sec = Win32::OLE->CreateObject("ADsSecurity");
  my $error_num;
  my $error_name;
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error creating security object (ADSSecurity) -> $error_num ($error_name)\n");
    if ($error_num eq "0x80004002") {
      _DebugComment("Error:  No such interface supported.\n  Note:  Make sure you have the ADSSecurity.DLL from the ADSI SDK regisered on this system\n");
    }
    return 0;
  }

  bless $new_mailbox,"Win32::OLE";
  
  my $sd = $sec->Get($new_mailbox->{msExchMailboxSecurityDescriptor});
  
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying security descriptor for mailbox -> $error_num ($error_name)\n");
    bless $new_mailbox,"Win32::Exchange";
    return 0;
  }
  my $dacl = $sd->{DiscretionaryAcl};
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying discretionary acl for mailbox -> $error_num ($error_name)\n");
    bless $new_mailbox,"Win32::Exchange";
    return 0;
  }

  foreach my $userid (@perms_list) {
    _DebugComment("      -Setting perms for $userid\n");
    my $ace = Win32::OLE->CreateObject("AccessControlEntry");
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error creating access control entry for mailbox -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }

    $ace->LetProperty('Trustee',$userid); 
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting trustee for mailbox -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }

    $ace->LetProperty('AccessMask',$ADS_RIGHT_EXCH_MODIFY_USER_ATT | $ADS_RIGHT_EXCH_MAIL_SEND_AS | $ADS_RIGHT_EXCH_MAIL_RECEIVE_AS);
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting access mask for mailbox -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }

    $ace->LetProperty('AceType', $ADS_ACETYPE_ACCESS_ALLOWED);
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting ace type for mailbox -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }
    $dacl->AddAce($ace);
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error adding access control entry to perms list -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }

    $sd->LetProperty("DiscretionaryAcl",$dacl); 
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting  discretionary acl on security security descriptor -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }
    $sec->SetSecurityDescriptor($sd);
    if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
      _DebugComment("error setting security descriptor on security object -> $error_num ($error_name)\n");
      bless $new_mailbox,"Win32::Exchange";
      return 0;
    }
  }
  $new_mailbox->SetInfo;
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting permissions on mailbox -> $error_num ($error_name)\n");
    bless $new_mailbox,"Win32::Exchange";
    return 0;
  }
  bless $new_mailbox,"Win32::Exchange";
  return 1;
}

sub AddDLMembers {
  my $ldap_provider;
  my $server_name;
  my $exch_dl_name;
  my @new_members;
  my $ou;
  my $org;
  if (scalar(@_) > 3) {
    $ldap_provider = $_[0];
    $server_name=$_[1];
    $exch_dl_name=$_[2];
    if (ref($_[3]) ne "ARRAY") {
      _DebugComment("members list must be an array reference\n");
      return 0;
    }
    @new_members=@{$_[3]};
    if (scalar(@_) == 4) {
      if ($ldap_provider->GetLDAPPath($server_name,$org,$ou)) {
        _DebugComment("returned -> o=$org,ou=$ou\n");
      } else {
        _DebugComment("Error Returning from GetLDAPPath\n");
        return 0;
      }
    } elsif (scalar(@_) == 6) {
      $org = $_[3];
      $ou = $_[4];
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
  my $exch_dl = $ldap_provider->GetObject("","LDAP://$server_name/cn=$exch_dl_name,cn=Distribution Lists,ou=$ou,o=$org");
  bless $ldap_provider,"Win32::Exchange";
  my $error_num;
  my $error_name;
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying distribution list ($exch_dl_name) -> $error_num ($error_name)\n");
    return 0;
  }
  
  my $exch_members = $exch_dl->{'member'}; #get the list
  if (ref($exch_members) eq "ARRAY") {
    _DebugComment("      -Array (2 or more members exist)\n");
  } else {
    _DebugComment("      -(Less than 2 members are named in this distribution list)\n");
    my $temp_exch_dl=$exch_members;
    undef ($exch_members);
    if ($temp_exch_dl) {
      _DebugComment("      -1 member exists\n");
      #So push the existing name to the Array
      push (@$exch_members, $temp_exch_dl);
    } else {
      _DebugComment("      -0 members exists\n");
    }
  }
  foreach my $username (@new_members) {
    _DebugComment("      -Adding $username to Distribution List: $exch_dl_name\n");
    my $duplicate;
    foreach my $dup (@$exch_members) {
      if (lc($dup) =~ lc("cn=$username,cn=Recipients,ou=$ou,o=$org")) {
        $duplicate = 1;
        last;
      }
    }
    if ($duplicate != 1) {
      push (@$exch_members, "cn=$username,cn=Recipients,ou=$ou,o=$org");
    }
  }
  $exch_dl->Put('member', $exch_members);
  $exch_dl->SetInfo;
  if (!_ErrorCheck(Win32::OLE->LastError(),"0x00000000",$error_num,$error_name)) {
    _DebugComment("error setting new member for distribution list ($exch_dl_name) -> $error_num ($error_name)\n");
    return 0;
  }
  return 1;
}

sub _ErrorCheck {
  my $last_ole_error = $_[0];
  my $last_error_expected = $_[1];
  my $error_num;
  my $error_name;
  $error_num = sprintf ("0x%08x",$last_ole_error);
  my @error_list = split(/\"/,$last_ole_error,3);
  $error_name = $error_list[1];
  if ($error_num ne $last_error_expected) {
    $_[2] = $error_num;
    $_[3] = $error_name;
    return 0;
  } else {
    return 1;
  }
}

sub _ReportArgError {
  _DebugComment("incorrect number of options passed to $_[0] ($_[1])\n");
  return 1;
}

sub _DebugComment {
  print "$_[0]" if ($DEBUG == 1);
  return 1;
}
1;
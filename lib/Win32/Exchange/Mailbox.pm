# Win32::Exchange::Mailbox
# Freely Distribute the code without modification.
#
# Creates and Modifies Exchange 5.5 and 2K Mailboxes
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

package Win32::Exchange::Mailbox;

use strict;
use vars qw ($VERSION $Version $DEBUG);

use Win32::OLE qw (in);
use Win32::OLE::Variant;
Win32::OLE->Initialize(Win32::OLE::COINIT_OLEINITIALIZE);
use Win32::Exchange::Const;

Win32::OLE->Option('_Unique' => 1);
#@ISA = qw(Win32::OLE);

my $Version;
my $VERSION = $Version = "0.041";
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
    if ($_[0] eq "Win32::Exchange::Mailbox") {
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

  my $class = "Win32::Exchange::Mailbox";
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
  if (Win32::Exchange::_AdodbExtendedSearch($server_name,"LDAP://$server_name","(&(objectClass=Computer)(rdn=$server_name))","rdn,distinguishedName",$result)) {
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
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  if ($type eq "IPerson") {
    #IPerson returns for CDO.Person (E2K)
    if ($mbx = _E2KCreateMailbox(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      bless $mbx,"Win32::Exchange::Mailbox";
      return $mbx;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($mbx = _E55CreateMailbox(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      bless $mbx,"Win32::Exchange::Mailbox";
      return $mbx;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
  return 0;
}

sub _E55CreateMailbox {
  my $ldap_provider;
  my $information_store_server;
  my $mailbox_alias_name;
  my $org = "";
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
      #placeholder
    } elsif (scalar(@_) == 4) {
      $container = $_[3];
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
  if ($org eq "") {
    if ($ldap_provider->GetLDAPPath($information_store_server,$org,$ou)) {
      _DebugComment("returned -> o=$org,ou=$ou\n",3);
    } else {
      _DebugComment("Error Returning from GetLDAPPath\n",1);
      return 0;
    }
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
  bless $ldap_provider,"Win32::Exchange::Mailbox";

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
  if (scalar(@_) > 3) {
    $provider = $_[0];
    $info_store_server = $_[1];
    $nt_pdc = $_[2];
    $mailbox_alias_name = $_[3];
    if (scalar(@_) == 4) {
      #placeholder..
    } elsif (scalar(@_) == 5) {
      $mailbox_ldap_path = $_[4]
    } elsif (scalar(@_) == 6) {
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
  Win32::Exchange::_StripBackslashes($nt_pdc,$pdc); 
  my $user_dist_name;
  if (!Win32::Exchange::_AdodbExtendedSearch($mailbox_alias_name,"LDAP://$pdc","(samAccountName=$mailbox_alias_name)","samAccountName,distinguishedName",$user_dist_name)) {
    _DebugComment("Error querying distinguished name for user in CreateMailbox (E2K)\n",1);
    return 0;
  }
 
  _DebugComment("user_dist_name = $user_dist_name\n",3);  
 
  bless $provider,"Win32::OLE";
  my $user_account = $provider->DataSource->Open("LDAP://$pdc/$user_dist_name",undef,adModeReadWrite);
  #   http://support.microsoft.com/default.aspx?scid=kb;EN-US;q321039

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
    if (!Win32::Exchange::LocateMailboxStore($info_store_server,$storage_group,$mb_store,$mailbox_ldap_path)) {
      return 0;
    }
  }
  _DebugComment("$mailbox_ldap_path\n",3);
  $info_store->CreateMailbox($mailbox_ldap_path);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed creating mailbox for $mailbox_alias_name ($error_num) $error_name\n",1);
    return 0;
  }
 
  $provider->DataSource->Save();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed saving mailbox for $mailbox_alias_name ($error_num) $error_name\n",1);
    return 0;
  }
  return $provider;
}

sub DeleteMailbox {
  my $error_num;
  my $error_name;
  my $provider = $_[0];

  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object for Exchange Server Determination for DeleteMailbox\n",1);
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  my $rtn = 0;
  if ($type eq "IPerson") {
    #IPerson returns for CDO.Person (E2K)
    if (_E2KDeleteMailbox(@_)) {
      $rtn = 1;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if (_E55DeleteMailbox(@_)) {
      $rtn = 1;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
  return $rtn;
}

sub _E2KDeleteMailbox {
  if (scalar(@_) != 1) {
    _ReportArgError("DeleteMailbox [E2K]",scalar(@_));
    return 0;
  }
  my $mailbox = $_[0];
  my $error_num;
  my $error_name;
  bless $mailbox,"Win32::OLE";
  my $interface = $mailbox->GetInterface("IMailboxStore");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error getting IMailboxStore interface for user mailbox deletion [E2K]\n",1);
    bless $mailbox,"Win32::Exchange::Mailbox";
    return 0;
  }
  $interface->DeleteMailbox();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error deleting user mailbox (DeleteMailbox) [E2K]\n",1);
    bless $mailbox,"Win32::Exchange::Mailbox";
    return 0;
  }

  $mailbox->Datasource->Save();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Error deleting user mailbox (Save) [E2K]\n",1);
    bless $mailbox,"Win32::Exchange::Mailbox";
    return 0;
  }
  _DebugComment("Mailbox deleted successfully",1);
  return 1;
}

sub _E55DeleteMailbox {
  my $ldap_provider;
  my $information_store_server;
  my $mailbox_alias_name;
  my $error_num;
  my $error_name;
  my $find_mb;
  if (scalar(@_) > 2) {
    $ldap_provider = $_[0];
    $information_store_server = $_[1];
    $mailbox_alias_name = $_[2];
  } else {
    _ReportArgError("DeleteMailbox [5.5] ",scalar(@_));
    return 0;
  }
  my $recipients_path;
  my $exch_mb_dn;
  my $path;
  
  my $Recipients = $ldap_provider->GetMailboxContainer($information_store_server,$mailbox_alias_name);
  bless $ldap_provider,"Win32::OLE";
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening recipients path on $information_store_server\n",1);
    return 0;
  }

  bless $Recipients,"Win32::OLE";
  my $original_ole_warn_value = $Win32::OLE::Warn;
  $Win32::OLE::Warn = 0; #Turn STDERR warnings off because we probably are going to get an error (0x80072030) if we are creating a new box.
  
  $Recipients->Delete("organizationalPerson", "cn=$mailbox_alias_name");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Unable to Get the mailbox object where the rdn == $mailbox_alias_name on $information_store_server ($error_num)\n",1);
    $Win32::OLE::Warn=$original_ole_warn_value;
    return 0;
  }
  $Win32::OLE::Warn=$original_ole_warn_value;
  return 1;
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
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  if ($type eq "IPerson") {
    #IPerson returns for CDO.Person (E2K)
    if ($mbx = _E2KGetMailbox(@_)) {
      bless $mbx,"Win32::Exchange::Mailbox";
      return $mbx;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($mbx = _E55GetMailbox(@_)) {
      bless $mbx,"Win32::Exchange::Mailbox";
      return $mbx;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
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
    if (Win32::Exchange::_AdodbExtendedSearch($mailbox_alias_name,"LDAP://$information_store_server","(&(objectClass=organizationalPerson)(rdn=$mailbox_alias_name))","rdn,distinguishedName",1,$exch_mb_dn)) {
      $recipients_path = "LDAP://$information_store_server/$exch_mb_dn";
    } else {
      _DebugComment("Error locating Exchange Mailbox on the server.\n",1);
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
    _DebugComment("Unable to Get the mailbox object where the rdn == $mailbox_alias_name on $information_store_server ($error_num)\n",1);
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
  Win32::Exchange::_StripBackslashes ($nt_dc,$dc);
  bless $provider,"Win32::OLE";
  my $user_dist_name;
  if (!Win32::Exchange::_AdodbExtendedSearch($mailbox_alias_name,"LDAP://$dc","(samAccountName=$mailbox_alias_name)","samAccountName,distinguishedName",$user_dist_name)) {
    _DebugComment("Error querying distinguished name for user in GetMailbox (E2K)\n",1);
    return 0;
  }
  $provider->DataSource->Open("LDAP://$dc/$user_dist_name",undef,adModeReadWrite);
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
  bless $user_obj,"Win32::OLE";
  if (_E2KIsMapiAware($dc,$mailbox_alias_name)) {
    _DebugComment("Error performing GetMailbox: user is not MAPI aware ($error_num)\n",2);
    bless $user_obj,"Win32::Exchange::Mailbox";
    return 0;
  } else {
    $provider->DataSource->Save();
    bless $user_obj,"Win32::Exchange::Mailbox";
    return $provider;
  }
}

sub GetUserObject {
  #Different from GetMailbox in that it gets the user object without a datasource->open.
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
  Win32::Exchange::_StripBackslashes ($nt_dc,$dc);
  
  my $user_dist_name;
  if (!Win32::Exchange::_AdodbExtendedSearch($mailbox_alias_name,"LDAP://$dc","(samAccountName=$mailbox_alias_name)","samAccountName,distinguishedName",$user_dist_name)) {
    _DebugComment("Error querying distinguished name for user in GetMailbox (E2K)\n",1);
    return 0;
  }

  my $user_obj;
  bless $provider,"Win32::OLE";
  $user_obj = Win32::OLE->GetObject("LDAP://$dc/$user_dist_name");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("Failed opening distinguishedname for GetUserObject ($error_num)\n",1);
    bless $user_obj,"Win32::Exchange::Mailbox";
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $user_obj,"Win32::Exchange::Mailbox";
  bless $provider,"Win32::Exchange::Mailbox";
  return $user_obj;
}

sub _E2KIsMapiAware {
  my $provider;
  my $mailbox_alias_name;
  my $nt_dc;
  my $dc;
  if (scalar(@_) == 2) {
    $nt_dc = $_[0];
    $mailbox_alias_name = $_[1];
  } else {
    _ReportArgError("IsMapiAware [E2K]",scalar(@_));
    return 0;
  }
  Win32::Exchange::_StripBackslashes ($nt_dc,$dc);
  
  my $mapi_recip;
  if (Win32::Exchange::_AdodbExtendedSearch($mailbox_alias_name,"LDAP://$dc","(&(samAccountName=$mailbox_alias_name)(mapirecipient=*))","samAccountName,mapirecipient",$mapi_recip)) {
    _DebugComment("This is not a MapiAware user account -- IsMapiAware (E2K)\n",3);
    return 1;
  } else {
    return 0;
  }
}

sub _E2KIsMailboxEnabled {
  my $provider;
  my $mailbox_alias_name;
  my $nt_dc;
  my $dc;
  if (scalar(@_) == 2) {
    $nt_dc = $_[0];
    $mailbox_alias_name = $_[1];
  } else {
    _ReportArgError("IsMailboxEnabled [E2K]",scalar(@_));
    return 0;
  }
  Win32::Exchange::_StripBackslashes ($nt_dc,$dc);
  
  my $user_dist_name;
  if (Win32::Exchange::_AdodbExtendedSearch($mailbox_alias_name,"LDAP://$dc","(samAccountName=$mailbox_alias_name)(mapirecipient=TRUE)","samAccountName,distinguishedName",$user_dist_name)) {
    return 1;
  } else {
    _DebugComment("This is not a MailboxEnabled user account -- IsMailboxEnabled (E2K)\n",3);
    return 0;
  }
}

sub _E2KIsMailEnabled {
  my $provider;
  my $mailbox_alias_name;
  my $nt_dc;
  my $dc;
  if (scalar(@_) == 2) {
    $nt_dc = $_[0];
    $mailbox_alias_name = $_[1];
  } else {
    _ReportArgError("IsMailEnabled [E2K]",scalar(@_));
    return 0;
  }
  Win32::Exchange::_StripBackslashes ($nt_dc,$dc);
  
  my $user_dist_name;
  if (Win32::Exchange::_AdodbExtendedSearch($mailbox_alias_name,"LDAP://$dc","(samAccountName=$mailbox_alias_name)(mapirecipient=FALSE)","samAccountName,distinguishedName",$user_dist_name)) {
    return 1;
  } else {
    _DebugComment("This is not a MailboxEnabled user account -- IsMailboxEnabled (E2K)\n",3);
    return 0;
  }
}

sub GetDLMembers {
  my $error_num;
  my $error_name;
  my $provider = $_[0];
  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object type for Exchange Server Determination during call to SetAttributes\n",1);
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KGetDLMembers(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55GetDLMembers(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
  return 0;
}

sub _E55GetDLMembers {
  my $ldap_provider;
  my $server_name;
  my $exch_dl_name;
  my @members;
  my $ou;
  my $org;
  my $find_dl;
  my $return_prop;
  if (scalar(@_) > 3) {
    $ldap_provider = $_[0];
    $server_name=$_[1];
    $exch_dl_name=$_[2];
    if (ref($_[3]) ne "ARRAY") {
      _DebugComment("members list must be an array reference\n",1);
      return 0;
    }
    if (scalar(@_) < 7) {
      if (scalar(@_) == 5) {
        $return_prop = $_[4];      
      }
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
      $org = $_[5];
      $ou = $_[6];
    } else {
      _ReportArgError("GetDLMembers [5.5]",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("GetDLMembers [5.5]",scalar(@_));
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
      if (Win32::Exchange::_AdodbExtendedSearch($exch_dl_name,"LDAP://$server_name","(&(objectClass=groupOfNames)(cn=$exch_dl_name))","cn,distinguishedName",$exch_dl_dn)) {
        $exch_dl_path = "LDAP://$server_name/$exch_dl_dn";
      } else {
        _DebugComment("Error locating Exchange DL on the server.  Member information unavailable.\n",1);
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
  
  if (ref($exch_dl->{Members}) eq "ARRAY") {
    _DebugComment("      -Array (2 or more members exist -- $return_prop)\n",3);
    if (lc($return_prop) ne 'distinguishedname') {
      foreach my $member (@{$exch_dl->{Members}}) {
        my $search_prop;
        if (!Win32::Exchange::_AdodbExtendedSearch($member,"LDAP://$server_name","(distinguishedName=$member)","distinguishedName,$return_prop",$search_prop)) {
          _DebugComment("Failed Adodb search for member property ($return_prop)\n",1);
          return 0;
        }
        push (@{$_[3]}, $search_prop);
      }
      return 1;
    } else {
      @{$_[3]} = @{$exch_dl->{Members}};
      return 1;

    }
  } else {    
    _DebugComment("      -Less than 2 members are named in this distribution list -- $return_prop\n",3);
    my $member = $exch_dl->{Members};
    if ($member) {
      _DebugComment("      -1 member exists -- $return_prop\n",3);
      if (lc($return_prop) ne 'distinguishedname') {
        my $search_prop;
        if (!Win32::Exchange::_AdodbExtendedSearch($member,"LDAP://RootDSE/dnsHostName","(distinguishedName=$member)","distinguishedName,$return_prop",$search_prop)) {
          _DebugComment("Failed Adodb search for member property ($return_prop)\n",1);
          return 0;
        }
        push (@{$_[3]}, $search_prop);
      } else {
        push (@{$_[3]}, $member);
      }
      return 1;
    } else {
      _DebugComment("      -0 members exists -- $return_prop\n",1);
      return 1;
    }
  }
}

sub _E2KGetDLMembers {
  my $error_num;
  my $error_name;
  my $group_dn;
  my $user_dn;
  my $provider = $_[0];
  my $group = $_[1];
  my $return_prop = 'distinguishedName';
  if (scalar(@_) != 3) {
    if (scalar(@_) == 4) {
      $return_prop = $_[3];
    } else {
      _ReportArgError("GetdDLMembers (E2K)",scalar(@_));
      return 0;
    }
  }
  if (ref($_[2]) ne "ARRAY") {
    _DebugComment("Third argument is the list of users you want return, and should be an array reference, but instead, it was a(an): ".ref($_[2])." reference\n",1);
    return 0;
  }

  if (!Win32::Exchange::_AdodbExtendedSearch($group,"LDAP://RootDSE/dnsHostName","(&(objectClass=group)(samAccountName=$group))","samAccountName,distinguishedName",$group_dn)) {
    _DebugComment("Failed Adodb search for dist list\n",1);
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
  if (ref($group_obj->{Member}) eq "ARRAY") {
    _DebugComment("      -Array (2 or more members exist -- $return_prop)\n",3);
    if (lc($return_prop) ne 'distinguishedname') {
      foreach my $member (@{$group_obj->{Member}}) {
        my $search_prop;
        if (!Win32::Exchange::_AdodbExtendedSearch($member,"LDAP://RootDSE/dnsHostName","(distinguishedName=$member)","distinguishedName,$return_prop",$search_prop)) {
          _DebugComment("Failed Adodb search for member property ($return_prop)\n",1);
          return 0;
        }
        push (@{$_[2]}, $search_prop);
      }
      return 1;
    } else {
      @{$_[2]} = @{$group_obj->{Member}};
      return 1;

    }
  } else {    
    _DebugComment("      -Less than 2 members are named in this distribution list -- $return_prop\n",3);
    my $member = $group_obj->{Member};
    if ($member) {
      _DebugComment("      -1 member exists -- $return_prop\n",3);
      if (lc($return_prop) ne 'distinguishedname') {
        my $search_prop;
        if (!Win32::Exchange::_AdodbExtendedSearch($member,"LDAP://RootDSE/dnsHostName","(distinguishedName=$member)","distinguishedName,$return_prop",$search_prop)) {
          _DebugComment("Failed Adodb search for member property ($return_prop)\n",1);
          return 0;
        }
        push (@{$_[2]}, $search_prop);
      } else {
        push (@{$_[2]}, $member);
      }
      return 1;
    } else {
      _DebugComment("      -0 members exists -- $return_prop\n",1);
      return 1;
    }
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
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KSetAttributes(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55SetAttributes(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
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
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    #no available support for this operation
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55GetOwner(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
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
    bless $new_mailbox,"Win32::Exchange::Mailbox";
    return 0;      
  }
  bless $new_mailbox,"Win32::Exchange::Mailbox";
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
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KSetPerms(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55SetPerms(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
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

sub MailEnable {
  my $error_num;
  my $error_name;
  my $provider = $_[0];
  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object type for Exchange Server Determination during call to MailEnable\n",1);
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KMailEnable(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55MailEnable(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
  return 0;
}

sub _E2KMailEnable {
  my $user_obj;
  my $interface;
  my $smtp_address = "";
  my $error_num;
  my $error_name;
  if (scalar(@_) > 0) {
    #The provider is useless at this point
    #$provider = $_[0];
    if (scalar(@_) == 3) {
      $user_obj = $_[1];
      $smtp_address = $_[2];
      if (!($smtp_address =~ /^smtp:.*/i && $smtp_address ne "")) {
        $smtp_address = "smtp:".$smtp_address;
      }
    } else {
      _ReportArgError("MailEnable (E2K)",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("MailEnable (E2K)",scalar(@_));
    return 0;
  }
  bless $user_obj,"Win32::OLE";
  
  my $dn = $user_obj->{ADsPath};
  $dn =~ /LDAP:\/\/(.*)\/.*/;
  my $dc = $1;
  if (_E2KIsMapiAware($dc,$user_obj->{samaccountname})) {
    _DebugComment("user account is already MAPI aware.  Cannot proceed with MailEnable (E2K)\n",0);
    return 0;  
  }

  $interface = $user_obj->GetInterface("IMailRecipient");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying mailbox interface during MailEnable (E2K) -> $error_num ($error_name)\n",1);
    return 0;
  }
  if ($smtp_address eq "") {
    $interface->MailEnable();
  } else {
    $interface->MailEnable($smtp_address);
  }
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error performing MailEnable on AD user object (E2K) -> $error_num ($error_name)\n",1);
    return 0;
    print "2\n";
  }

  $user_obj->Datasource->Save();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error performing Save on AD User object during MailEnable (E2K) -> $error_num ($error_name)\n",1);
    return 0;
  }
  if ($smtp_address eq "") {
    return 1;
  } else {
    if (ref($user_obj->{proxyaddresses}) eq "ARRAY") {
      _DebugComment("      -Array (2 or more addresses exist)\n",3);
      my @proxy_addr = @{$user_obj->{proxyaddresses}};
      my $found_addr = 0;
      foreach my $proxyaddr (@proxy_addr) {
        if (lc($proxyaddr) eq lc($smtp_address)) {
          $found_addr = 1;
        }
      }
      if ($found_addr == 0) {
        push (@proxy_addr,$smtp_address);
        $user_obj->{proxyaddresses} = @proxy_addr;
      }
      return 1;
    } else {    
      _DebugComment("      -Less than 2 addresses exist\n",3);
      my $proxyaddresses = $user_obj->{proxyaddresses};
      if ($proxyaddresses) {
        _DebugComment("      -1 address exists\n",3);
        my @proxy_addr;
        push (@proxy_addr,$user_obj->{proxyaddresses});
        if (lc($user_obj->{proxyaddresses}) ne lc($smtp_address)) {
          push (@proxy_addr,$smtp_address);
        }
        $user_obj->{proxyaddresses} = @proxy_addr;
        return 1;
      } else {
        _DebugComment("      -0 addresses exists\n",3);
        $user_obj->{proxyaddresses}= $smtp_address;
        return 1;
      }
    }
  }
}

sub _E55MailEnable {
  #sorry, no function yet.
  return 0;
}

sub MailDisable {
  my $error_num;
  my $error_name;
  my $provider = $_[0];
  bless $provider,"Win32::OLE";
  Win32::OLE->LastError(0);
  my $type = Win32::OLE->QueryObjectType($provider);
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("failed querying OLE Object type for Exchange Server Determination during call to MailEnable\n",1);
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KMailDisable(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55MailDisable(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
  return 0;
}

sub _E2KMailDisable {
  my $user_obj;
  my $interface;
  my $smtp_address = "";
  my $error_num;
  my $error_name;
  if (scalar(@_) > 0) {
    #The provider is useless at this point
    #$provider = $_[0];
    if (scalar(@_) == 2) {
      $user_obj = $_[1];
    } else {
      _ReportArgError("MailDisable (E2K)",scalar(@_));
      return 0;
    }
  } else {
    _ReportArgError("MailDisable (E2K)",scalar(@_));
    return 0;
  }
  bless $user_obj,"Win32::OLE";

  my $dn = $user_obj->{ADsPath};
  $dn =~ /LDAP:\/\/(.*)\/.*/;
  my $dc = $1;
  if (!_E2KIsMailEnabled($dc,$user_obj->{samaccountname})) {
    _DebugComment("user account is not MailEnabled.  Cannot proceed with MailDisable (E2K)\n",0);
    return 0;  
  }

  $interface = $user_obj->GetInterface("IMailRecipient");
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error querying mailbox interface during MailDisable (E2K) -> $error_num ($error_name)\n",1);
    return 0;
  }

  $interface->MailDisable();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error performing MailDisable on AD user object (E2K) -> $error_num ($error_name)\n",1);
    return 0;
  }

  $user_obj->Datasource->Save();
  if (!ErrorCheck("0x00000000",$error_num,$error_name)) {
    _DebugComment("error performing Save on AD User object during MailDisable (E2K) -> $error_num ($error_name)\n",1);
    return 0;
  } else {
    return 1;
  }
}

sub _E55MailDisable {
  #sorry, no function yet.
  return 0;
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
    bless $provider,"Win32::Exchange::Mailbox";
    return 0;
  }
  bless $provider,"Win32::Exchange::Mailbox";
  
  my $rtn;
  if ($type eq "IPerson") {
    #IPerson returns should CDO.Person (E2K)
    if ($rtn = _E2KAddDLMembers(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  } else {
    #nothing returns for ADsNamespaces (E5.5)
    if ($rtn = _E55AddDLMembers(@_)) {
      bless $provider,"Win32::Exchange::Mailbox";
      return $rtn;
    }
  }
  bless $provider,"Win32::Exchange::Mailbox";
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
      if (Win32::Exchange::_AdodbExtendedSearch($exch_dl_name,"LDAP://$server_name","(&(objectClass=groupOfNames)(cn=$exch_dl_name))","cn,distinguishedName",$exch_dl_dn)) {
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
        if (!Win32::Exchange::_AdodbExtendedSearch($username,"LDAP://$server_name","(&(objectClass=organizationalPerson)(rdn=$username))","rdn,distinguishedName",$exch_mb_dn)) {
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

  if (!Win32::Exchange::_AdodbExtendedSearch($group,"LDAP://RootDSE/dnsHostName","(&(objectClass=group)(samAccountName=$group))","samAccountName,distinguishedName",$group_dn)) {
    _DebugComment("Failed Adodb search for dist list\n",1);
    return 0;
  }

  foreach my $username (@user_list) {
    _DebugComment("Adding $username to $group\n",4);
    if (!Win32::Exchange::_AdodbExtendedSearch($username,"LDAP://RootDSE/dnsHostName","(&(objectClass=user)(samAccountName=$username))","samAccountName,distinguishedName",$user_dn)) {
      _DebugComment("Failed Adodb search for user: $username\n",1);
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

sub _ReportArgError {
  my $rtn = Win32::Exchange::_ReportArgError($_[0],$_[1]);
  return $rtn;
}

sub _DebugComment {
  my $rtn = Win32::Exchange::_DebugComment($_[0],$_[1],$DEBUG);
  return $rtn;
}

sub ErrorCheck {
  my $rtn = Win32::Exchange::ErrorCheck($_[0],$_[1],$_[2]);
  return $rtn;
}

1;
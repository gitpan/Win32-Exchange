use Win32::Exchange;
use Win32::AdminMisc;
#if you don't have/use AdminMisc, you can get it by typing:
#
#ppm install Win32-AdminMisc --location=http://www.roth.net/perl/packages
#
#from the command line of any ActivePerl-enabled PC.
#
#or just set $domain and $pdc to your DOMAIN and \\PDC

$domain = Win32::DomainName();
$pdc = Win32::AdminMisc::GetPDC($domain);
$mailbox_alias_name='thisisatest';
$mailbox_full_name="This $mailbox_alias_name Isatest";
$info_store_server="HOMEEXCH2";
$mta_server=$info_store_server; #this could be different, but for testing, we'll set them the same

if (!Win32::Exchange::GetVersion($info_store_server,\%ver) ) {
  die "$rtn - Error returning into main from GetVersion\n";
}

print "version      = $ver{ver}\n";
print "build        = $ver{build}\n";
print "service pack = $ver{sp}\n";
if (!($provider = Win32::Exchange->new($ver{'ver'}))) {
  die "$rtn - Error returning into main from new ($Win32::Exchange::VERSION)\n";
}

my @PermsUsers;
push (@PermsUsers,"$domain\\$mailbox_alias_name");
push (@PermsUsers,"$domain\\Exchange Perm Users"); #Group that needs perms to the mailbox...

if ($ver{ver} eq "5.5") {
  if (!Win32::Exchange::GetLDAPPath($info_store_server,$org,$ou)) {
    print "Error returning into main from GetLDAPPath\n";
    exit 1;
  }
  print "GetLDAPPath succeeded\n";
  if ($mailbox = $provider->GetMailbox($info_store_server,$mailbox_alias_name,$org,$ou)) {
    print "Mailbox already existed\n";
    if ($mailbox->SetOwner("$domain\\$mailbox_alias_name")) {
      print "SetOwner in GetMailbox worked!\n";
    }
    if ($mailbox->SetPerms(\@PermsUsers)) {
      print "Successfully set perms in GetMailbox\n";  
    } else {
      die "Error setting perms from GetMailbox\n";  
    }
  } else {
    $mailbox = $provider->CreateMailbox($info_store_server,$mailbox_alias_name,$org,$ou);
    if (!$mailbox) {
      die "error creating mailbox\n";
    }
    print "We created a mailbox!\n";
    if ($mailbox->SetOwner("$domain\\$mailbox_alias_name")) {
      print "SetOwner worked\n";  
    } else {
      print "SetOwner failed\n";  
    }
    if ($mailbox->GetOwner($nt_user,0x2)) {
      print "GetOwner worked: owner = $nt_user\n";  
    } else {
      print "GetOwner failed\n";  
    }

    $mailbox->GetPerms(\@array);
    
    foreach my $acl (@array) {
      print "   trustee - $acl->{Trustee}\n";  
      print "accessmask - $acl->{AccessMask}\n";  
      print "   acetype - $acl->{AceType}\n";  
      print "  aceflags - $acl->{AceFlags}\n";  
      print "     flags - $acl->{Flags}\n";  
      print "   objtype - $acl->{ObjectType}\n";  
      print "inhobjtype - $acl->{InheritedObjectType}\n";  
    }

    if ($mailbox->SetPerms(\@PermsUsers)) {
      print "Successfully set perms\n";  
    } else {
      die "Error setting perms\n";  
    }
  }
  
  #$Exchange_Info{'Deliv-Cont-Length'}='6000'; 
  #$Exchange_Info{'Submission-Cont-Length'}='6000'; 
  $Exchange_Info{'givenName'}="This";
  $Exchange_Info{'sn'}="Isatest";
  $Exchange_Info{'cn'}=$mailbox_full_name;
  $Exchange_Info{'mail'}="$mailbox_alias_name\@insight.com";
  $Exchange_Info{'rfc822Mailbox'}="$mailbox_alias_name\@insight.com"; 
  #You can add any attributes to this hash that you can set via exchange for a mailbox

  #$rfax="RFAX:$Exchange_Info{'cn'}\@"; #this can set the Rightfax SMTP name for Exchange-enabled Rightfax mail delivery
  #push (@$Other_MBX,$rfax);

  $smtp="smtp:another_name_to_send_to\@insight.com"; 
  push (@$Other_MBX,$smtp);
  #be careful with 'otherMailbox'es..  You are deleting any addresses that may exist already
  #if you set them via 'otherMailbox' and don't get them first (you are now forewarned).
  $Exchange_Info{'otherMailbox'}=$Other_MBX;

  if (!Win32::Exchange::GetDistinguishedName($mta_server,"Home-MTA",$Exchange_Info{"Home-MTA"})) {
    print "Failed getting distinguished name for Home-MTA on $info_store_server\n";
    exit 0;
  }
  if (!Win32::Exchange::GetDistinguishedName($info_store_server,"Home-MDB",$Exchange_Info{"Home-MDB"})) {
    print "Failed getting distinguished name for Home-MDB on $info_store_server\n";
    exit 0;
  }

  if ($mailbox->SetAttributes(\%Exchange_Info)) {
    print "SetAttributes worked\n";  
  } else {
    print "SetAttributes failed\n";  
  }

  my @new_dl_members;
  push (@new_dl_members,$mailbox_alias_name);
  $provider->AddDLMembers($info_store_server,"newdltest",\@new_dl_members); 

} elsif ($ver{ver} eq "6.0") {
  $storage_group = ""; #you'd need to define this if you had more than 1 storage group on 1 server.
  $mailbox_store = ""; #you'd need to define this if you had more than 1 mailbox store on 1 or more storage groups.
  if (Win32::Exchange::LocateMailboxStore($info_store_server,$storage_group,$mailbox_store,$store_name,\@counts)) {
    print "storage group = $storage_group\n";
    print "mailbox store = $mailbox_store\n";
    print "located store distinguished name= $store_name\n";
    print "$info_store_server\n";
    print "  Total:\n";
    print "    storage groups = $counts[0]\n";
    print "    mailbox stores = $counts[1]\n";
  }
  if ($mailbox = $provider->GetMailbox($pdc,$mailbox_alias_name)) {
    print "Got Mailbox successfully\n";
  } else {
    print "Mailbox did not exist\n";
    if ($mailbox = $provider->CreateMailbox($info_store_server,
                                            $pdc,
                                            $mailbox_alias_name,
                                            "insight.com"
                                           )
       ) {
      print "Mailbox create succeeded\n";
    } else {
      die "Failure is the option that you have selected!\n";
    }
    
  }
  #be careful with proxy addresses..  You are deleting any addresses that may exist already
  #if you set them via ProxyAddresses (you are now forewarned).
  push (@$proxies,'SMTP:'.$mailbox_alias_name.'@manross.net');
  push (@$proxies,'smtp:secondary@manross.net');
  push (@$proxies,'smtp:primary@manross.net');
  push (@$proxies,'smtp:tertiary@manross.net');

  $Attributes{"IMailRecipient"}{ProxyAddresses} = $proxies;
  
  #  $Attributes{"ExchangeInterfaceName"}{Property} = value; #with this method you should be able to set any value
  #                                                           imaginable.....

  $Attributes{"IMailRecipient"}{IncomingLimit} = 6000;
  $Attributes{"IMailRecipient"}{OutgoingLimit} = 6000;
  $Attributes{"IMailboxStore"}{EnableStoreDefaults} = 0;
  $Attributes{"IMailboxStore"}{StoreQuota} = 100; #at 100KB starts getting warnings
  $Attributes{"IMailboxStore"}{OverQuotaLimit} = 120; #at 120KB can't send...  I THINK...
  $Attributes{"IMailboxStore"}{HardLimit} = 130; #at 130KB, can't do anything...  I THINK...
  if (!$mailbox->SetAttributes(\%Attributes)) {
    die "Error setting 2K Attributes\n";
  } else {
    print "Set Attributes correctly\n";
  }

  my @PermUsers;
  push (@PermUsers,"$domain\\$mailbox_alias_name");
  push (@PermUsers,"$domain\\Exchange Perms Admin"); #Group that needs perms to the mailbox...

  if (!$mailbox->SetPerms(\@PermUsers)) {
    die "Error setting 2K Perms\n";
  } else {
    print "Set 2K Perms correctly\n";
  }
  my @new_dl_members;
  push (@new_dl_members,$mailbox_alias_name);
  if ($provider->AddDLMembers("_homelist",\@new_dl_members)) {
    print "Add successful to DL\n";
  } else {
    die "Error adding distlist member\n";
  }
  exit 1;
}
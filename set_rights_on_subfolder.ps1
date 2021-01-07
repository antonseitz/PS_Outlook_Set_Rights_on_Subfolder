
param(

[string] $owner,
[string] $user,
[string] $folder,
[string] $recursive

)




function set-outlookacl {

 # Check if user already in ACL 
        $acl=get-mailboxfolderpermission $fname -user $user  


        if( $acl)
            {
            "User in ACL alreay present!"
            $acl
            " 
            "
            "Should we delete this Entry?"

            $input = read-host "Enter ENTER to confirm deletetion
Enter other key to continue without deleting acl-entry "
            if($input -eq "")
                {
                remove-mailboxfolderpermission $fname -user $user -confirm:$false
            }
        }

        Add-MailboxFolderPermission $fname -user $user -AccessRights owner
}


"


DON'T FORGET TO CONNECT TO YOUR TENANT WITH

Import-Module exchangeonlinemanagement"
"Connect-ExchangeOnline

Otherwise, get-mailboxfolderpermission will fail!


"

$input=Read-host "
Type any key and hit ENTER to connect to Exchange Online now
Just hit Enter to continue without connecting"


if ($input ){


Import-Module exchangeonlinemanagement
Connect-ExchangeOnline
}

# USAGE
if (! $folder -and ! $owner ){ 
	write-host("Usage: "  +  $MyInvocation.MyCommand.Name + "  -owner mailbox_owner  [ -folder subfolder ] [ -user username ] ")
	write-host(" -owner username_of_mailbox_owner  ")
	write-host(" -folder subfolder in Outlook on which user should get access ] ")
    " Examplse: \Posteingang    \Posteingang\Subfolder \Inbox\Subfolder "
    write-host(" -user Username of user, who should get permissions] ")
    write-host(" -recursive $true : set acl entry for all subfolder of folder")
	exit
}

if ( $folder -and $owner) { 

"this are current permissions on `"" + $folder +  "`" :"  

get-mailboxfolderpermission ($owner + ":" + $folder)
}


$folder_with_slash=$folder.Replace("\","/")

if( $user -and $owner -and $folder -and $folder_with_slash){


#$folders=Get-MailboxFolderStatistics $owner




# recursive
if($recursive){
    foreach ( $f in ( Get-MailboxFolderStatistics $owner | where {$_.FolderPath.contains($folder_with_slash) -eq $true } )) 
        {#$f; 
        $fname = $owner + ":" + $f.FolderPath.replace("/","\");
        ""
        "Ordner"
        $fname; 
        set-outlookacl 

       
    }
}


else  
# not recursive but single path
    {$fname= $owner + ":" + $folder;
    set-outlookacl 

}
}


else{


"
Use -owner and -user to give rights on this folder to a certain user!"


}


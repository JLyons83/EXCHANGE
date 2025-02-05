#before running this script need to connect to exchange - connect-exchangeonline and also to connect to microsoft graph - "Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All"


#function to timestamp logs
function Get-TimeStamp {
    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    
}


#ous for getting users to search if user is synced with 365
$ou1 = ""
$ou2 = ""
$ou3 = ""
$ou4 = ""
$ou5 = ""



#get users from groups fixed above
$usersinAD = @(
    Get-ADUser -Filter * -SearchBase $ou1
    Get-ADUser -Filter * -SearchBase $ou2
    Get-ADUser -Filter * -SearchBase $ou3
    Get-ADUser -Filter * -SearchBase $ou4
    Get-ADUser -Filter * -SearchBase $ou5
)



$users = Import-Csv "C:\Users\ajeremy\Documents\bulkmailusers-test.csv"


foreach($user in $users){

$checkcontactexists = get-contact $user.MicrosoftOnlineServicesID -erroraction silentlycontinue

if($checkcontactexists){
Write-output "$(Get-TimeStamp) contact $($user.MicrosoftOnlineServicesID) exists, deleting...." | Out-file C:\admin\log\newmailuser2.txt -append
remove-mailcontact $user.MicrosoftOnlineServicesID -erroraction silentlycontinue
}
else{}


$userin365 = $usersinAD | Where-Object { $_.SamAccountName -eq $user.username } -ErrorAction SilentlyContinue

if($userin365){
Write-output "$(Get-TimeStamp) user $($user.username) exists in 365 and is syncing from AD, starting process to set-aduser" | Out-file C:\admin\log\newmailuser2.txt -append
#set domains
$domain = "@jct.ac.il"
$domain_offsite = "@jctacil.onmicrosoft.com"

$username = $user.username

#set proxyaddresses variable for attribute
$proxyaddress1 = "SMTP:$username$domain"
$proxyaddress2 = "smtp:$username$domain_offsite"
$Proxyaddresses="$proxyaddress1,$proxyaddress2"

#set mail variable for attribute
$mail = "$Username$domain"

Write-output "$(Get-TimeStamp) user $($user.username)  starting process to set-aduser" | Out-file C:\admin\log\newmailuser2.txt -append
#sets attributes in AD for exchange 365 email
Set-ADUser -Identity $username -Add @{proxyaddresses=$Proxyaddresses;mailnickname=$username} -Replace @{mail=$mail;UserPrincipalName=$mail}
#adds to group for exchange email
add-adgroupmember -identity "exchange users" -members $username

start-sleep -seconds 2

Write-output "$(Get-TimeStamp) user $($user.username)  syncing changes with 365" | Out-file C:\admin\log\newmailuser2.txt -append
do {

    # Try to start AD Sync
  try{  $sync = Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction SilentlyContinue }
  catch{}
    # Check the result
    if ($sync.Result -eq "Success") {
        Write-output "$(Get-TimeStamp) AD Sync completed successfully! for user $($user.username)" | Out-file C:\admin\log\newmailuser2.txt -append
        break  # Exit the loop
    } else {
        Write-output "$(Get-TimeStamp) Sync not successful, retrying in 10 seconds..." | Out-file C:\admin\log\newmailuser2.txt -append
        Start-Sleep -Seconds 10
    }

} until ($sync.Result -eq "Success")

start-sleep -seconds 210


do {


try{
#creates variable to check if forwarding rule is created
$forwardingrule = (get-mailbox $user.MicrosoftOnlineServicesID -ErrorAction SilentlyContinue).ForwardingSmtpAddress 

#creates forwarding rule to external email verifying that emails are not saved in mailbox
                set-mailbox -Identity $user.MicrosoftOnlineServicesID -ForwardingSmtpAddress $user.externalemailaddress1 -DeliverToMailboxAndForward $false}
                catch{}

#checks if forwarding rule is created
if($forwardingrule -eq "smtp:$($user.externalemailaddress1)"){
    write-output "$(Get-TimeStamp) forwarding rule for mailbox $($user.MicrosoftOnlineServicesID) to $($user.externalemailaddress1) was created successfully" | Out-file C:\admin\log\newmailuser2.txt -append
    break #if created exists loop
}
else{

        Write-output "$(Get-TimeStamp) mailbox for $($user.MicrosoftOnlineServicesID) is still being created" | Out-file C:\admin\log\newmailuser2.txt -append
        Start-Sleep -Seconds 10
}


}until($forwardingrule -eq "smtp:$($user.externalemailaddress1)")


}

else{

$checkmailuser = get-mailuser -identity $user.externalemailaddress1 -ErrorAction SilentlyContinue

if($checkmailuser){
}
else{

     #check if user has user in 365 and then add mailbox with forwarding rule
     try{

     #gets user in 365 and creates variable
        $365user = get-mguser -userid $user.MicrosoftOnlineServicesID -erroraction SilentlyContinue

         #checks if 365 user exists
            if($365user){



            #log 365 user exists
            Write-output "$(Get-TimeStamp) 365 user $($user.MicrosoftOnlineServicesID) exists" | Out-file C:\admin\log\newmailuser2.txt -append


            #if exists applies license to user - which creates the mailbox
                Set-MgUserLicense -userid $user.MicrosoftOnlineServicesID -AddLicenses @{SkuId = "78e66a63-337a-4a9a-8959-41c6654dfb56"} -RemoveLicenses @() -erroraction SilentlyContinue
            
            #log creating mailbox
            Write-output "$(Get-TimeStamp) creating mailbox for $($user.MicrosoftOnlineServicesID)" | Out-file C:\admin\log\newmailuser2.txt -append

            #waits 50 seconds to finish creating mailbox
                    start-sleep -Seconds 70
            
            #log creating forwarding rule
            Write-output "$(Get-TimeStamp) creating forwarding rule to $($user.externalemailaddress1) from mailbox $($user.MicrosoftOnlineServicesID)" | Out-file C:\admin\log\newmailuser2.txt -append

            #creates forwarding rule to external email verifying that emails are not saved in mailbox
                do {


                        try{
                        #creates variable to check if forwarding rule is created
                        $forwardingrule = (get-mailbox $user.MicrosoftOnlineServicesID -ErrorAction SilentlyContinue).ForwardingSmtpAddress 

                        #creates forwarding rule to external email verifying that emails are not saved in mailbox
                                set-mailbox -Identity $user.MicrosoftOnlineServicesID -ForwardingSmtpAddress $user.externalemailaddress1 -DeliverToMailboxAndForward $false}
                                catch{}

                        #checks if forwarding rule is created
                        if($forwardingrule -eq "smtp:$($user.externalemailaddress1)"){
                            write-output "forwarding rule for mailbox $($user.MicrosoftOnlineServicesID) to $($user.externalemailaddress1) was created successfully" | Out-file C:\admin\log\newmailuser2.txt -append
                            break #if created exists loop
                        }
                        else{

                        Write-output "mailbox for $($user.MicrosoftOnlineServicesID) is still being created" | Out-file C:\admin\log\newmailuser2.txt -append
                        Start-Sleep -Seconds 10
                             }


                }until($forwardingrule -eq "smtp:$($user.externalemailaddress1)")

                $createdmailbox = get-mailbox $user.MicrosoftOnlineServicesID -ErrorAction SilentlyContinue
            
            #log mailbox created successfully and forwarding rule
                if($createdmailbox){
                Write-output "$(Get-TimeStamp) mailbox for $($user.MicrosoftOnlineServicesID) created successfully" | Out-file C:\admin\log\newmailuser2.txt -append
                }
                else{
                Write-output "$(Get-TimeStamp) mailbox creation for $($user.MicrosoftOnlineServicesID) failed" | Out-file C:\admin\log\newmailuser2.txt -append
                }
                if($createdmailbox.ForwardingSmtpAddress -eq "smtp:$($user.externalemailaddress1)"){
                Write-output "$(Get-TimeStamp) forwarding rule to $($user.externalemailaddress1) for $($user.MicrosoftOnlineServicesID) created successfully" | Out-file C:\admin\log\newmailuser2.txt -append
                }
                else{
                 Write-output "$(Get-TimeStamp) forwarding rule to $($user.externalemailaddress1) for $($user.MicrosoftOnlineServicesID) failed"  | Out-file C:\admin\log\newmailuser2.txt -append
                }


                }

     else{
#process to create maiuser - if not user in 365 then create mailuser
        
        #log starting mailuser creation
          Write-output "$(Get-TimeStamp) 365 user $($user.MicrosoftOnlineServicesID) doesnt exist proceeding with process to create mail user" | Out-file C:\admin\log\newmailuser2.txt -append  

           

    #delete emailaddresses from distribution list
       
        foreach($column in $user.PSObject.Properties.Name) {

             # Check if the value is equal to 1 and therefore delete from distribution list

            if ($user.$column -eq 1) {

            #log deleting member of distribution list
                Write-output "$(Get-TimeStamp) deleting $($user.externalemailaddress1) from distribution list $column" | Out-file C:\admin\log\newmailuser2.txt -append

                Remove-DistributionGroupMember -identity $column -member $user.externalemailaddress1 -confirm:$false
                         }

                }

        #delete email address from address book


        start-sleep 1
        

            # Loop through each column
        foreach ($column in $user.PSObject.Properties.Name){
            # Check if the value is $true
            if ($user.$column -eq 1) {

                        #log deleting member of distribution list
                Write-output "$(Get-TimeStamp) deleting $($user.externalemailaddress1) from contacts" | Out-file C:\admin\log\newmailuser2.txt -append
           
                remove-mailcontact -identity $($user.externalemailaddress1) -confirm:$false

                #breaks loop once finds 1 column equal to 1
                break
                }
            }

                #sleep to guarantee deletion of contact before creation of mail user
     start-sleep -Seconds 5

            #log verify that mailcontact was deleted or not
                    foreach ($column in $user.PSObject.Properties.Name){
            # Check if the value is $true
            if ($user.$column -eq 1) {
            $mailcontact = get-mailcontact -identity $user.externalemailaddress1 -ErrorAction SilentlyContinue 
                if($mailcontact){
                Write-output "$(Get-TimeStamp) deletion of $($user.externalemailaddress1) from contacts failed" | Out-file C:\admin\log\newmailuser2.txt -append
                }
                else{
                Write-output "$(Get-TimeStamp) deletion of $($user.externalemailaddress1) from contacts was successful" | Out-file C:\admin\log\newmailuser2.txt -append
                                }
                        }
                }
                
    


    #log mailuser creating
      Write-output "$(Get-TimeStamp) creating mail user for $($user.externalemailaddress1)" | Out-file C:\admin\log\newmailuser2.txt -append


    #Randomly Chooses Password
    Add-Type -AssemblyName 'System.Web'
    $minLength = 20 ## characters
    $maxLength = 40 ## characters
    $length = Get-Random -Minimum $minLength -Maximum $maxLength
    $nonAlphaChars = 5
    $password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
    $secPw = ConvertTo-SecureString -String $password -AsPlainText -Force

    #create mail user
               $mailuser = New-MailUser -Name $user.Name -ExternalEmailAddress `
             $user.ExternalEmailAddress1 -MicrosoftOnlineServicesID $user.MicrosoftOnlineServicesID `
             -Password $secPw
             
             
             
   #start sleep until process of creating user finishes

   start-sleep -Seconds 20
 

        #log verify mailuser creation
         if($mailuser){
          Write-output "$(Get-TimeStamp) creation of mail user for $($user.externalemailaddress1) was successful" | Out-file C:\admin\log\newmailuser2.txt -append
         }
         else{
         Write-output "$(Get-TimeStamp) creation of mail user for $($user.externalemailaddress1) failed" | Out-file C:\admin\log\newmailuser2.txt -append
         }

  #add mailuser to distribution list
 

        
    foreach($column in $user.PSObject.Properties.Name) {

        # Check if the value is equal to 1 and therefore add in distribution list

        if ($user.$column -eq 1) {
          Write-output "$(Get-TimeStamp) adding mail user  $($user.externalemailaddress1) to distribution list $column" | Out-file C:\admin\log\newmailuser2.txt -append 
            add-DistributionGroupMember -identity $column -member $user.externalemailaddress1 -confirm:$false
        }
    }

 
 }


      

     }
     catch{}
     }
     }
     }
 
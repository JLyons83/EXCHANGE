$users = Import-Csv "C:\Users\ajeremy\Documents\bulkmailusers-test.csv"


foreach($user in $users){



    #delete emailaddresses from distribution list
       
        foreach($column in $user.PSObject.Properties.Name) {

             # Check if the value is equal to 1 and therefore delete from distribution list

            if ($user.$column -eq 1) {
                Write-host "deleting $user.externalemailaddress1 from distribution list $column"
                Remove-DistributionGroupMember -identity $column -member $user.externalemailaddress1 -confirm:$false
                         }

                }

        #delete email address from address book


        start-sleep 1
        

            # Loop through each column
        foreach ($column in $user.PSObject.Properties.Name){
            # Check if the value is $true
            if ($user.$column -eq 1) {
            Write-host "deleting $user.externalemailaddress1 from contacts"
                remove-mailcontact -identity $($user.externalemailaddress1) -confirm:$false

                #breaks loop once finds 1 column equal to 1
                break
                }
            }
    
    #sleep to guarantee deletion of contact before creation of mail user
    start-sleep -Seconds 2

     #check if user has user in 365 and then add mailbox with forwarding rule
     try{

        $365user = get-mguser -userid $user.MicrosoftOnlineServicesID -erroraction SilentlyContinue

            if($365user){

                Set-MgUserLicense -userid $user.MicrosoftOnlineServicesID -AddLicenses @{SkuId = "78e66a63-337a-4a9a-8959-41c6654dfb56"} -RemoveLicenses @()

                    start-sleep -Seconds 50

                set-mailbox -Identity $user.MicrosoftOnlineServicesID -ForwardingSmtpAddress $user.externalemailaddress1 -DeliverToMailboxAndForward $false
                }
     else{
           
      #if not user in 365 then create mailuser

      
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

   start-sleep -Seconds 5
 


  #add mailuser to distribution list
 

        
    foreach($column in $user.PSObject.Properties.Name) {

        # Check if the value is equal to 1 and therefore add in distribution list

        if ($user.$column -eq 1) {
          
            add-DistributionGroupMember -identity $column -member $user.externalemailaddress1 -confirm:$false
        }
    }

 
 }


      

     }
     catch{}
     }
 
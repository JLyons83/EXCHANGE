$users = Import-Csv "C:\Users\ajeremy\Documents\bulkmailusers-test.csv"


foreach($user in $users){



    #delete emailaddresses from distribution list


   
        foreach($column in $user.PSObject.Properties.Name) {

             # Check if the value is equal to 1 and therefore delete from distribution list

            if ($user.$column -eq 1) {
                Write-Output "Column '$column' in row '$($user.name)' is TRUE, performing action..."
                Remove-DistributionGroupMember -identity $column -member $user.externalemailaddress1 -confirm:$false
                         }

                }

        #delete email address from address book


        start-sleep 1
        

            # Loop through each column
        foreach ($column in $user.PSObject.Properties.Name){
            # Check if the value is $true
            if ($user.$column -eq 1) {

                remove-mailcontact -identity $($user.externalemailaddress1) -confirm:$false

                #breaks loop once finds 1 column equal to 1
                break
                }
            }
    
    #sleep to guarantee deletion of contact before creation of mail user
    start-sleep -Seconds 2

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
 
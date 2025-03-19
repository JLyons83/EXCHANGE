add-type -name user32 -namespace win32 -memberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'
$consoleHandle = (get-process -id $pid).mainWindowHandle

# load your form or whatever ...

# hide console
[win32.user32]::showWindow($consoleHandle, 0)

set-executionpolicy bypass -Scope Process

Install ExchangeOnlineManagement module if not installed
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Force -AllowClobber
}


Import-Module ExchangeOnlineManagement

connect-exchangeonline 



# Create a GUI input form
Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form



$form.Text = "Update External Email Address"
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = "CenterScreen"

$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "MailUser Email Address:"
$label1.Location = New-Object System.Drawing.Point(10,20)
$form.Controls.Add($label1)

$textbox1 = New-Object System.Windows.Forms.TextBox
$textbox1.Location = New-Object System.Drawing.Point(10,40)
$textbox1.Width = 260
$form.Controls.Add($textbox1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Text = "New External Email Address:"
$label2.Location = New-Object System.Drawing.Point(10,70)
$form.Controls.Add($label2)

$textbox2 = New-Object System.Windows.Forms.TextBox
$textbox2.Location = New-Object System.Drawing.Point(10,90)
$textbox2.Width = 260
$form.Controls.Add($textbox2)

$button = New-Object System.Windows.Forms.Button
$button.Text = "Update"
$button.Location = New-Object System.Drawing.Point(100,130)
$button.Add_Click({
    $mailUser = $textbox1.Text
    $newExternalEmail = $textbox2.Text
    
    if ($mailUser -and $newExternalEmail) {
        try {
            Set-MailUser -Identity $mailUser -ExternalEmailAddress $newExternalEmail
            [System.Windows.Forms.MessageBox]::Show("External Email Updated Successfully", "Success")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter all fields", "Error")
    }
})

$form.Controls.Add($button)
$form.Topmost = $true
$form.ShowDialog()


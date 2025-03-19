#צירך לבדוק שהמיילים תקינים - המייל הראשון מבוסס על ההנחה שהמייל הוא השם משתמש בלי קידומת דהיינו השתי אותיות הראשונות
# $MAIL.TO ולהזין ידנית המייל המבוקש במשתנה $EMAIL אם לא, אפשר להוריד את המשתנה
 
$user = $env:Username 

$email = $user.Substring(2) + "@moin.gov.il"

$outlook = New-Object -ComObject Outlook.Application

$Mail = $Outlook.CreateItem(0)

$Mail.To = $email

# כאן תרשום מייל חיצוני חשוב להגדיר את תיבת המייל לשלוח מייל אוטומטי בחזרה
$Mail.CC = "testemail@gmail.com"

#כאן מגדירים את הנושא
 
$Mail.Subject = "Test Mail"

$Mail.Send()

 # לפי עומס ברשת אולי צריך להגדיל את הזמן 
#  שמחכה עד שזה סורק את אאוטלוק למיילים שנשלחו

Sleep 60

$outlook = New-Object -ComObject Outlook.Application

$olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type]

$namespace = $outlook.GetNameSpace(“MAPI”)

$folder = $namespace.getDefaultFolder($olFolders::olFolderInBox)

#הנושאים של המיילים שהמערכת סורק צריך להיות זהים עם הנושא שנקבע למעלה
#קודם זה מחפש את המייל שנשלח למייל פנימי ואז המייל שנשלח אוטומטי מהמייל החיצוני
$emailsArrived = $folder.items | Where-Object {($_.Subject -eq "Test mail") -or ($_.Subject -eq 'Re: Test Mail');}

$EmailsToDelete = $emailsArrived.Count

#בודק כאן עם יש רק שתי מיילים שהגיע ונמצא בתוך המערכת
#echo $EmailsToDelete

If($EmailsToDelete -eq 2){
Write-Host "Mail Sent Successfully" -ForegroundColor Green;
$emailsArrived.Delete()
}
else{
    Write-Host "Mail Test Failed" -ForegroundColor Red
}

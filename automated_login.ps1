 #Create an Internet Explorer object
$ie = New-Object -COMObject "InternetExplorer.Application"
$ie.Visible= $true # Make it visible
#$ie.GetType().FullName
#while ($ie.Busy -eq $true) {Start-Sleep -Seconds 3;} 
#$Ie|Get-Member
# Open all websites
if ($Null -eq $ie) {
    $ie
}
$ie.Navigate("http://facebook.com/")

#Navigate2() function to open in new tab in the same IE window
$ie.Navigate2("http://https://github.com/",0x1000) 
#$ie.Navigate2("http://outlook.com",0x1000) 
#$ie.Navigate2("https://www.linkedin.com/",0x1000) 

# Script to wait till webpage is downloaded into the browsers.
while ($ie.Busy -eq $true) {Start-Sleep -Seconds 3;} 
#Do {Start-Sleep 10} While ($ie.busy)


# Facebook Login
# To Identify the Facebook window among the existing IE tabs.
$ie = (New-Object -COM "Shell.Application").Windows() |Where-Object{$_.locationname -like '*Facebook*'}
If($null -eq $ie){ $ie.Refresh()}
while ($ie.Busy -eq $true){Start-Sleep -seconds 3}
if ($Null -eq $ie.Document) { $ie.Document }
# Feed in your credentials to input fields on the web page
$usernamefield = $ie.Document.getElementById("email")
$usernamefield.value = "mail"
$passwordfield = $ie.document.getElementById("pass")
$passwordfield.value = "pass"
$Link=$ie.Document.getElementById("loginbutton") | Select-Object -First 1
$Link.click()

# Outlook Login
# To Identify the Outlook window among the existing IE tabs.
#$ie = (New-Object -COM "Shell.Application").Windows() | Where-Object{$_.locationname -like '*Outlook*'}
#If($null -eq $ie.Document){ $ie.Refresh()}
#while ($ie.Busy -eq $true){Start-Sleep -seconds 1; Write-Output 'loading outlook  tab...'}

# Feed in your credentials to input fields on the web page
#$usernamefield = $ie.Document.getElementsByClassName("input__field input__field--with-label")
#$usernamefield.value = "mail"
#$passwordfield = $ie.Document.getElementByID("userpassword")
#$passwordfield.value = "pass"
#$Link=$ie.Document.getElementsByClassName("btn-primary loginButton")
#$Link.click()

# GitHub Login
# To Identify the GitHub window among the existing IE tabs.
$ie = (New-Object -COM "Shell.Application").Windows() | Where-Object{$_.locationname -like '*GitHub*'}
If($null -eq $ie.Document){ $ie.Refresh()}
while ($ie.Busy -eq $true){Start-Sleep -seconds 1; Write-Output 'loading Live.com '}

# Feed in your credentials to input fields on the web page
$usernamefield = $ie.Document.getElementByID('login_field')
$usernamefield.value = 'mail'
$passwordfield = $ie.Document.getElementByID('password')
$passwordfield.value = 'pass'
$Link=$ie.Document.getElementsByClassName('btn btn-primary btn-block')
$Link.click() 

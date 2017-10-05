Function CreateSignature  # Here we acutally create the signature
{
    # First we start by making the appropriate registry changes. In the environment, there should only be Office 2016, however it could be remotely possible that 
    # a few remote machines were missed in the upgrade, so we also include Office 2013

    if (test-path "HKCU:\Software\Microsoft\Office\15.0\Common\General") {
    get-item -path HKCU:\Software\Microsoft\Office\15.0\Common\General | new-Itemproperty -name NewSignature -value CompanySignature -propertytype string -force
    }

    if (test-path "HKCU:\Software\Microsoft\Office\15.0\Common\General") {
    get-item -path HKCU:\Software\Microsoft\Office\15.0\Common\General | new-Itemproperty -name ReplySignature -value CompanySignature -propertytype string -force
    }

    if (test-path "HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook\Setup") {
    remove-itemproperty "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\setup" -Name "First-Run"
    }

    if (test-path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\MailSettings") {
    get-item -path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\MailSettings|new-Itemproperty -name NewSignature -Value signature -propertyType ExpandString -Force
    }

    if (test-path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\MailSettings") {
    get-item -path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\MailSettings|new-Itemproperty -name ReplySignature -Value signature -propertyType ExpandString -Force
    }

    if (test-path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\Mail") {
    get-item -path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\Mail|new-Itemproperty -name "Send Pictures With Document" -Value 1 -propertyType DWORD -Force
    }

    if (test-path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Setup") {
    remove-itemproperty "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\setup" -Name "First-Run"
    }

# Create the directory for the template file and the signature file to be stored
$FolderLocation = $UserDataPath + '\\Microsoft\\CompanySignature'  
mkdir $FolderLocation -force 

# Copy the template file to the directory on the local computer. This is done because we dont want to accidently
# make changes to the template file. Plus it prevents file contention.
copy-item -Path $Logonserver\CompanySignature\SigTemplate1.htm -Destination $Folderlocation\SigTemplate1.htm 

# All we need to do is replace the appropriate strings in the template file with the appropriate variables from AD
# and create the signature.htm file

(Get-Content $Folderlocation\SigTemplate1.htm) | Foreach-Object {
    $_ -replace '%%DisplayName%%', $strName `
       -replace '%%Title%%', $strTitle `
       -replace '%%PhoneNumber%%', $strphone `
       -replace '%%PagerNumber%%', $strextension `
       -replace '%%Email%%', $stremail `
    } | Set-Content $FolderLocation\\signature.htm

}

# First of all we gather a bunch of necessary information from AD about the user

$strName = $env:username
$UserDataPath = $Env:appdata
$Logonserver = $env:LOGONSERVER +"\netlogon"
$strFilter = "(&(objectCategory=User)(samAccountName=$strName))"
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter
$objPath = $objSearcher.FindOne()
$objUser = $objPath.GetDirectoryEntry()
$strName = $objUser.FullName
$strTitle = $objUser.Title
$strPhone = $objUser.telephoneNumber
$strextension = $objuser.pager
$strEmail = $objUser.mail

$FolderLocation = $UserDataPath + '\\Microsoft\\CompanySignature'

# If this file does not exist, we are going to create the new signature. 
If ( -not (test-path $FolderLocation\\SigTemplate1.htm)){
    CreateSignature
    }
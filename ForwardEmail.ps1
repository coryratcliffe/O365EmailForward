#Log into Microsoft 365
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

$Again = "Y"
while($Again-eq"Y")
{
    #Prompts for the Name of the person whose email is being forwarded
    $FName = Read-Host 'Whose email?'
    $FName = ($FName)

    #Prompts for the name of who the email is being forwarded too
    $ForwardedToFName = Read-Host 'Forwarded to?'

    #Finds the Email address of the person who the eamil is being forwarded too
    Get-Mailbox | select DisplayName,UserPrincipalName| where {$_.DisplayName -eq $ForwardedToFName} | Export-csv C:\Users\cratcliffe\Desktop\TestGettingEmail.csv -NoTypeInformation
    $user = Import-csv C:\Users\cratcliffe\Desktop\TestGettingEmail.csv
    $ForwardEmailAddress = $user."UserPrincipalName"

    #Applies the users email forwarding
    Write-Host ("Forwarding " +$FName +" email to " + $ForwardEmailAddress)
    Set-Mailbox $FName -ForwardingAddress $ForwardEmailAddress -DeliverToMailboxAndForward $True
    Write-Host (" ")
    Write-Host ("Completed Eamil Forward")
    Write-Host (" ")
    $Again = Read-Host "Would you like to do another email forward (Y/N)?"

    #Removes Created CSV file
    Remove-Item C:\Users\cratcliffe\Desktop\TestGettingEmail.csv
}
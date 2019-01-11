﻿Import-Module -Name 'ActiveDirectory'
. "$PSScriptRoot\Config.ps1"
$adParams = @{
    Filter = $Script:Config.Filter
    Properties = @('UserPrincipalName','Title','Department','msDS-cloudExtensionAttribute14','Surname','GivenName')
}
$selectProps = @(
    @{Name = 'Anvandare'  ;Expression = {$_.'msDS-cloudExtensionAttribute14'}}
    @{Name = 'Epost'      ;Expression = {$_.UserPrincipalName}}
    @{Name = 'Fornamn'    ;Expression = {$_.GivenName}}
    @{Name = 'Efternamn'  ;Expression = {$_.Surname}}
    @{Name = 'Forvaltning';Expression = {$_.Department}}
    @{Name = 'Befattning' ;Expression = {$_.Title}}
)
# Cannot use Export-Csv since InfoCaption requires that header columns are not quoted.
'Anvandare,Epost,Fornamn,Efternamn,Forvaltning,Befattning' |
    Out-File "$PSScriptRoot\kungsbacka.csv" -Encoding UTF8 -Force
Get-ADUser @adParams | Select-Object -Property $selectProps | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 |
    Out-File "$PSScriptRoot\kungsbacka.csv" -Encoding UTF8 -Append
# Decrypt password
$password = (New-Object -TypeName 'PSCredential' -ArgumentList @('not used', ($Script:Config.Password | ConvertTo-SecureString))).GetNetworkCredential().Password
# -k ignores certificate errors and allows for self signed certificates
# without trusting them first. It's ok since this should be a well known FTP server.
& "$PSScriptRoot\curl.exe" -k -s --ssl -T "$PSScriptRoot\kungsbacka.csv" "ftp://$($Script:Config.Host)" -u "$($Script:Config.Username):$password" | Out-Null
Remove-Item -Path "$PSScriptRoot\kungsbacka.csv"

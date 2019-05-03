﻿Import-Module -Name 'ActiveDirectory'
. "$PSScriptRoot\Config.ps1"
$adParams = @{
    Filter = $Script:Config.Filter
    Properties = @('UserPrincipalName','Title','Department','msDS-cloudExtensionAttribute14','Surname','GivenName','Manager','EmailAddress')
}
# Cannot use Export-Csv since InfoCaption requires that header columns are not quoted.
'Anvandare,Epost,Fornamn,Efternamn,Forvaltning,Befattning,Eposttva,Chefsnamn,Chefepost' |
    Out-File "$PSScriptRoot\kungsbacka.csv" -Encoding UTF8 -Force
Get-ADUser @adParams |
    Foreach-Object -Process {
        $o = [PSCustomObject]@{
            Anvandare = $_.'msDS-cloudExtensionAttribute14'
            Epost = $_.UserPrincipalName
            Fornamn = $_.GivenName
            Efternamn = $_.Surname
            Forvaltning = $_.Department
            Befattningq = $_.Title
            Eposttva = $_.EmailAddress
            Chefsnamn = ''
            Chefepost = ''
        }
        if ($_.Manager) {
            $m = Get-ADUser $_.Manager -Properties 'DisplayName','EmailAddress' -ErrorAction SilentlyContinue
            if ($m) {
                $o.Chefsnamn = $m.DisplayName
                $o.Chefepost = $m.EmailAddress
            }
        }
        $o
    } | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File "$PSScriptRoot\kungsbacka.csv" -Encoding UTF8 -Append
# Decrypt password
$password = (New-Object -TypeName 'PSCredential' -ArgumentList @('not used', ($Script:Config.Password | ConvertTo-SecureString))).GetNetworkCredential().Password
# -k ignores certificate errors and allows for self signed certificates without trusting them first.
& "$PSScriptRoot\curl.exe" -k -s --ssl -T "$PSScriptRoot\kungsbacka.csv" "ftp://$($Script:Config.Host)" -u "$($Script:Config.Username):$password" | Out-Null
Remove-Item -Path "$PSScriptRoot\kungsbacka.csv"

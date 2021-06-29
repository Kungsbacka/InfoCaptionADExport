Import-Module -Name 'ActiveDirectory'
. "$PSScriptRoot\Config.ps1"
$adParams = @{
    Filter = $Script:Config.Filter
    Properties = @('UserPrincipalName','Title','Department','msDS-cloudExtensionAttribute14','Surname','GivenName','Manager','EmailAddress','PhysicalDeliveryOfficeName')
}
# Cannot use Export-Csv since InfoCaption requires that header columns are not quoted.
'Anvandare,Epost,Fornamn,Efternamn,Forvaltning,Enhet,Befattning,Eposttva,Chefsnamn,Chefepost' |
    Out-File "$PSScriptRoot\$($Script:Config.Filename)" -Encoding UTF8 -Force
Get-ADUser @adParams |
    Foreach-Object -Process {
        $o = [PSCustomObject]@{
            Anvandare = $_.'msDS-cloudExtensionAttribute14'
            Epost = $_.UserPrincipalName
            Fornamn = $_.GivenName
            Efternamn = $_.Surname
            Forvaltning = $_.Department
            Enhet = $_.PhysicalDeliveryOfficeName
            Befattning = $_.Title
            Eposttva = $_.EmailAddress
            Chefsnamn = ''
            Chefepost = ''
        }
        if ($_.Manager) {
            $m = Get-ADUser $_.Manager -Properties 'DisplayName','msDS-cloudExtensionAttribute14' -ErrorAction SilentlyContinue
            if ($m) {
                $o.Chefsnamn = $m.DisplayName
                $o.Chefepost = $m.'msDS-cloudExtensionAttribute14'
            }
        }
        $o
    } | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Select-Object -Skip 1 | Out-File "$PSScriptRoot\$($Script:Config.Filename)" -Encoding UTF8 -Append
# Decrypt password
$password = (New-Object -TypeName 'PSCredential' -ArgumentList @('not used', ($Script:Config.Password | ConvertTo-SecureString))).GetNetworkCredential().Password
# -k ignores certificate errors and allows for self signed certificates without trusting them first.
& "$PSScriptRoot\curl.exe" -k -s --ssl -T "$PSScriptRoot\$($Script:Config.Filename)" "ftp://$($Script:Config.Host)" -u "$($Script:Config.Username):$password" | Out-Null
Remove-Item -Path "$PSScriptRoot\$($Script:Config.Filename)"

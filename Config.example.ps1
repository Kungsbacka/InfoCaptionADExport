$Config = @{
    Host = '<FTP server>'
    Username = ''
    # Encrypted password.
    # To encrypt:
    #   Start a PowerShell window as the user that is going to run the script.
    #   Run: (Get-Credential).Password | ConvertFrom-SecureString | Set-Clipboard
    #   Enter a username (not used, but cannot be empty) and the password
    #   Paste the result into this file
    Password = ''
    # Filter used by Get-ADUser
    Filter = ''
}
# Check MFA methods registered for users from a CSV file (by UPN) and store results in an object
# Author: Brandon Scarberry
# Date: 2025-05-20

# Connect to Microsoft Graph (if not already connected)
Connect-MgGraph -Scopes "UserAuthenticationMethod.Read.All", "User.Read.All"

# Import the CSV file with a 'UserPrincipalName' column
$users = Import-Csv -Path "C:\users\BrandonScarberry\Downloads\exportUsers_2025-5-20.csv"

# Initialize array to hold results
$userMfaInfo = @()

foreach ($user in $users) {
    $upn = $user.UserPrincipalName

    # Get user details
    $userDetails = Get-MgUser -UserId $upn -Property DisplayName

    # Get MFA methods for the user
    $methods = Get-MgUserAuthenticationMethod -UserId $upn

    # Collect method types
    $mfaTypes = @()
    foreach ($method in $methods) {
        # The '@odata.type' property indicates the method type
        $type = $method.AdditionalProperties.'@odata.type'
        switch ($type) {
            '#microsoft.graph.passwordAuthenticationMethod' { $mfaTypes += 'Password' }
            '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' { $mfaTypes += 'Microsoft Authenticator' }
            '#microsoft.graph.phoneAuthenticationMethod' { $mfaTypes += "Phone ($($method.PhoneType))" }
            '#microsoft.graph.fido2AuthenticationMethod' { $mfaTypes += 'FIDO2' }
            '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' { $mfaTypes += 'Windows Hello for Business' }
            '#microsoft.graph.emailAuthenticationMethod' { $mfaTypes += 'Email' }
            default { $mfaTypes += $type }
        }
    }

    # Store in object
    $userMfaInfo += [PSCustomObject]@{
        UserPrincipalName = $upn
        DisplayName       = $userDetails.DisplayName
        MfaMethods        = $mfaTypes -join ', '
    }
}

# Output the results
$userMfaInfo
function Get-eDOCSLogon {
    <#
    .SYNOPSIS
    Logs a user onto eDOCS and generates a security token that is required for
    all other actions via the API.
    
    .DESCRIPTION
    All actions with the eDOCS API require passing a security token to eDOCS.
    This function performs a logon to eDOCS and returns a hashtable indicating
    if the login succeeded and if so the security token associated with the
    login.
    
    .PARAMETER dmUser
    Username of the account to login to eDOCS DM with.
    
    .PARAMETER dmPass
    eDOCS DM password for the specified account.
    
    .PARAMETER dmLibrary
    eDOCS DM library to be logged into typically EDOCS for prod or EDOCS_TEST
    for test.
    
    .EXAMPLE
    $login = Get-eDOCSLogon -dmUser "currentuser" -dmPass "secrets" -dmLibrary "EDOCS_TEST"

    Would return @{
        "success" = $true;
        "dst" = "12345abcdefgh";
    }

    If the logon succeeded or @{
        "success" = $false;
        "dst" = -1;
    }
    #>

    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $dmUser,
        [parameter(Mandatory = $true, Position = 1)]
        [String]
        $dmPass,
        [parameter(Mandatory = $true, Position = 2)]
        [String]
        $dmLibrary
    )

    # Login to the given eDOCS library. More details about this process available on 
    # page 55 of the eDOCS DM 5.3.1 DM API Reference Guide.
    $login = New-Object -ComObject PCDClient.PCDLogin
    $login.AddLogin(0, $dmLibrary, $dmUser, $dmPass)
    $login.Execute()
    $dmDST = $login.GetDST()

    if ($dmDST.Length -eq 0) {
        'Failed to login to eDOCS ' + $dmUser + '; Error: ' + $($login.ErrDescription) | Out-Host
        @{
            success = $false
            dst     = 0
        }
    } else {
        @{
            success = $true
            dst     = $dmDST
        }
    }
}

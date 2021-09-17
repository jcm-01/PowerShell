<#
    .AUTHOR
        JC Mocke (https://github.com/jcm-repo/PowerShell/)
    
    .DESCRIPTION
        Report on installed license keypacks and issued licenses.
    
    .USAGE
        Run, or schedule, the report on a Remote Desktop license server.
    
    .PARAMETERS
        -MailFrom
            Specify the email address from which the email notifications will be sent.

        -MailTo
            Specify the email address to which the report will be sent.

        -MailSubject
            Specify the email subject.

        -MailServer
            Specify an IP address or DNS name for your SMTP server.
#>

#region Set Initial Variables

    [CmdletBinding()]
    Param
    (
        [string] $MailFrom = 'RD Licensing <alerts@acme-corp.com>',

        [string] $MailTo = 'alerts@acme-corp.com',

        [string] $MailSubject = 'Remote Desktop Licensing Report',

        [string] $MailServer = 'smtp.acme-corp.net',

        [switch] $Test,

        [string] $TestAddress = 'admin.name@acme-corp.com'
    )

    $HTMLStylePath = "$( $PSScriptRoot )\HTMLStyling.css"

    # Send email to alternate address during testing (PowerShell ISE detected)
    If ($psISE -or $Test.IsPresent )
    {
        $MailTo      = $TestAddress
        $MailSubject = "$( $MailSubject ) (Test)"
    }

#endregion

#region Import Cascading Style Sheet
    
    If ( Test-Path -Path $HTMLStylePath )
    {
        $HTMLStyle = Get-Content -Path $HTMLStylePath
    }
    Else
    {
        Write-Warning "The CSS file located at $( $HTMLStylePath ) was not found!"
        
        $HTMLStyle = ''
    }

#endregion

#region Get License Information

    $LicenseKeyPacks = Get-WmiObject -Class 'Win32_TSLicenseKeyPack'

    $IssuedLicenses = Get-WmiObject -Class 'Win32_TSIssuedLicense' |
        Sort-Object -Property @(
            'sIssuedToUser'
        )

#endregion

#region Create License Version Array (Needed for lookups)
    
    $LicenseVersions = @{}
    
    $LicenseKeyPacks |
        Select-Object -Property @(
            'KeyPackId'
            'ProductVersion'
        ) |
        ForEach-Object -Process {
            $LicenseVersions[$_.KeyPackId] = $_.ProductVersion
        }

#endregion

#region Process License KeyPack Information
    
    $HTMLParamters = @{
        Fragment   = $true
        PreContent = "<h1>License KeyPack(s) on $( $env:COMPUTERNAME ) ($( ( $LicenseKeyPacks | Measure ).Count )):</h1>"
    }
    
    $LicenseKeyPacksHTML = $LicenseKeyPacks |
        Select-Object -Property @(
            @{
                Name       = 'KeyPack ID'
                Expression = { $_.KeyPackId }
            }
            @{
                Name       = 'Product Version'
                Expression = { $_.ProductVersion }
            }
            @{
                Name       = 'Type and Model'
                Expression = { $_.TypeAndModel }
            }
            @{
                Name       = 'Total'
                Expression = { $_.TotalLicenses }
            }
            @{
                Name       = 'Issued'
                 Expression = { $_.IssuedLicenses }
            }
            @{
                Name       = 'Available'
                Expression = { $_.AvailableLicenses }
            }
            @{
                Name       = 'Expiration Date'
                Expression = { $_.ConvertToDateTime( $_.ExpirationDate ).ToString( 'dd MMM yyyy @ HH:mm' ) }
            }
        ) |
        ConvertTo-Html @HTMLParamters

#endregion

#region Process Issuance Information
    
    $IssuedLicenses = $IssuedLicenses |
        Select-Object -Property @(
            @{
                Name       = 'Issued To'
                Expression = { $_.sIssuedToUser.Split( '\' )[1].ToUpper() }
            }
            @{
                Name       = 'Product Version'
                Expression = { $LicenseVersions[$_.KeyPackID] }
            }
            @{
                Name       = 'Issue Date'
                Expression = { $_.ConvertToDateTime( $_.IssueDate ).ToString( 'dd MMM yyyy @ HH:mm' ) }
            }
            @{
                Name       = 'Expiration Date'
                Expression = { $_.ConvertToDateTime( $_.ExpirationDate ).ToString( 'dd MMM yyyy @ HH:mm' ) }
            }
        )

#endregion

#region Prepare Email Body
    
    $HTMLParamters = @{
        Fragment   = $true
        PreContent = "<h1>Issued Licenses ($( ( $IssuedLicenses | Measure ).Count )):</h1>"
    }

    $IssuedLicensesHTML = $IssuedLicenses |
        ConvertTo-Html @HTMLParamters
    
    $MailBody = $HTMLStyle + $LicenseKeyPacksHTML + $IssuedLicensesHTML |
        Out-String

#endregion

#region Send Email
    
    $SendMailParameters = @{
        From       = $MailFrom
        To         = $MailTo
        Subject    = $MailSubject
        Body       = $MailBody
        BodyAsHtml = $true
        SmtpServer = $MailServer
    }

    Send-MailMessage @SendMailParameters

#endregion

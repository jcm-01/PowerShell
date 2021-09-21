
<#
    .AUTHOR
        JC Mocke

    .DESCRIPTION
        Find expiring/expired server certificates by checking all servers.

    .USAGE
        Run, or schedule, the report from any domain member on which RSAT is installed. The user account under which
        the script is executed will need access to the remote system.

    .PARAMETERS
        -IssuerName
            Specify the certificate issuer name to filter out certificates not issued by your PKI.
        
        -ExpirationThreshold
            Specify the amount of days to look ahead for expiring certificates.

        -MailFrom
            Specify the email address from which the report will be sent.
        
        -MailTo
            Specify the email address to which the report will be sent.
        
        -MailSubject
            Specify the email subject.
        
        -MailServer
            Specify an IP address or DNS name for your SMTP server.
        
        -Test
            The report will be executed as a test which overrides the "MailTo" email address.

        -TestAddress
            Specify the email address to which the test report will be sent.
#>

#region Set Initial Variables

    [CmdletBinding()]
    Param
    (
        [string] $IssuerName = 'acme-corp',
        
        [int] $ExpirationThreshold = 100,

        [string] $MailFrom = 'Server Certificates <alerts@acme-corp.com>',
        
        [string] $MailTo = 'alerts@acme-corp.com',
        
        [string] $MailSubject = 'Expiring Server Certificates',
        
        [string] $MailServer = 'smtp.acme-corp.net',

        [switch] $Test,

        [string] $TestAddress = 'admin.name@acme-corp.com'
    )
    
    $DateFormat = 'ddd, dd MMM yyyy, HH:mm:ss'
    
    $HTMLStylePath = "$( $PSScriptRoot )\HTMLStyling.css"

    If ($psISE -or $Test.IsPresent )
    {
        $MailTo      = $TestAddress
        $MailSubject = "$( $MailSubject ) (Test)"
    }

#endregion

#region Set Initial Functions

    Function GetRemainingDays
    {
        $TimeSpanParams = @{
            Start = Get-Date
            End   = $_.NotAfter
        }
    
        "$( ( New-TimeSpan @TimeSpanParams ).Days ) days"
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

#region Get AD Servers
    
    $ServerParams = @{
        Filter     = '*'
        Properties = 'OperatingSystem'
    }

    $Servers = Get-ADComputer @ServerParams |
        Where-Object -FilterScript {
            $_.DistinguishedName -match 'Servers|Domain Controllers' -or
            $_.OperatingSystem -match 'Server'
        } |
        Sort-Object -Property @(
            'Name'
        )

#endregion

#region Get Certificate Information

    $Certificates = ForEach ( $Server in $Servers.DNSHostName )
    {
        # Since the normal name property might be limited to 15 characters, the DNS hostname is used instead.
        $Hostname = $Server.Split( '.' )[0].ToUpper()
    
        $ConnectionTestParams = @{
            ComputerName = $Server
            Count        = 1
            Quiet        = $true
        }
    
        If ( $Hostname -eq $env:COMPUTERNAME )
        {
            Get-ChildItem -Path 'Cert:\LocalMachine\My' |
                Select-Object -Property @(
                    'Thumbprint'
                    'Issuer'
                    'NotBefore'
                    'NotAfter'
                    'DnsNameList'
                    'Subject'
                )
        }
        ElseIf ( Test-Connection @ConnectionTestParams )
        {
            $WinRMTestParams = @{
                ComputerName = $Server
                ErrorAction  = 'SilentlyContinue'
            }
            
            If ( Test-WSMan @WinRMTestParams )
            {
                Write-Host "Checking $( $Hostname )..."
            
                $CommandParam = @{
                    ComputerName = $Server
                    ScriptBlock  = {
                        Get-ChildItem -Path 'Cert:\LocalMachine\My' |
                            Select-Object -Property @(
                                'Thumbprint'
                                'Issuer'
                                'NotBefore'
                                'NotAfter'
                                'Subject'
                            )
                    }
                }
    
                Invoke-Command @CommandParam
            }
            Else
            {
                Write-Warning "$( $Hostname ) failed WinRM test!"
            }
        }
        Else
        {
            Write-Warning "$( $Hostname ) not available!"
        }
    }

#endregion

#region Process Certificate Information

    $Certificates = $Certificates |
        Where-Object -FilterScript {
            $_.Issuer -match $IssuerName -and
            $_.NotAfter -le ( Get-Date ).AddDays( $ExpirationThreshold )
        } |
        Select-Object -Property @(
           @{
                Name       = 'Server Name'
                Expression = { $_.PSComputerName.Split( '.' )[0].ToUpper() }
           }
           @{
                Name       = 'Issuer'
                Expression = { $_.Issuer.Split( ',' )[0].Split( '=' )[1] }
           }
           @{
                Name       = 'Start Date'
                Expression = { $_.NotBefore.ToString( $DateFormat ) }
           }
           @{
                Name       = 'Expiration Date'
                Expression = { $_.NotAfter.ToString( $DateFormat ) }
           }
           @{
                Name       = 'Expiration Countdown'
                Expression = { GetRemainingDays }
           }
           'Thumbprint'
        ) |
        Sort-Object -Property @(
            'Expiration Date'
            'Server Name'
        )

#endregion

#region Prepare Email Body
    
    $HTMLParams = @{
        Fragment   = $true
        PreContent = "<h1>The following certificates will expire within $( $ExpirationThreshold ) days ($( $Certificates.Thumbprint.Count )):</h1>"
    }

    $CertificatesHTML = $Certificates |
        ConvertTo-Html @HTMLParams
    
    $MailBody = $HTMLStyle + $CertificatesHTML |
        Out-String

#endregion

#region Send Email
    
    If ( $Certificates )
    {
        $SendMailParameters = @{
            From       = $MailFrom
            To         = $MailTo
            Subject    = $MailSubject
            Body       = $MailBody
            BodyAsHtml = $true
            SmtpServer = $MailServer
        }

        Send-MailMessage @SendMailParameters
    }

#endregion

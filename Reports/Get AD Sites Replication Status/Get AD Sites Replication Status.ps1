<#
    .AUTHOR
        JC Mocke (https://github.com/jcm-repo/PowerShell/)
    
    .DESCRIPTION
        Gathers Active Directory replication partner metadata from each domain controller and sends a report. 
    
    .USAGE
        Run, or schedule, the report from any domain member on which RSAT is installed.
    
    .PARAMETERS
        -MailFrom
            Specify the email address from which the report will be sent.
        
        -MailTo
            Specify the email address to which the report will be sent.
        
        -MailSubject
            Specify the email subject.
        
        -MailServer
            Specify the hostname or IP address of your SMTP server.
        
        -Test
            The report will be executed as a test - this replaces the "MailTo" address with the "TestAddress" and
            append "(Test)" to the email subject.

            Note, executing the report from PowerShell ISE will automatically execute the report as a test.

        -TestAddress
            Specify the email address to which the test report will be sent.
#>

#region Set Initial Variables

    [CmdletBinding()]
    Param
    (
        [string] $MailFrom = 'AD Sites & Services <alerts@acme-corp.com>',

        [string] $MailTo = 'alerts@acme-corp.com',

        [string] $MailSubject = 'AD Replication Report',

        [string] $MailServer = 'smtp.acme-corp.net',

        [switch] $Test,

        [string] $TestAddress = 'admin.name@acme-corp.com'
    )
    
    # Send email to alternate address during testing (PowerShell ISE detected)
    If ($psISE -or $Test.IsPresent )
    {
        $MailTo      = $TestAddress
        $MailSubject = "$( $MailSubject ) (Test)"
    }

    $HTMLStylePath = "$( $PSScriptRoot )\HTMLStyling.css"

    $DateFormat = 'dd MMM, HH:mm'

    $Domain = ( Get-ADForest ).RootDomain

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

#region Set Initial Functions
    
    Function Resolve-StatusCode
    {
        Switch ($_.LastReplicationResult)
        {
            '0'     { '0: No errors' }
            '1127'  { '1127: While accessing the hard disk, a disk operation failed even after retries' }
            '1256'  { '1256: The remote system is not available' }
            '1396'  { '1396: Logon Failure: The target account name is incorrect' }
            '1722'  { '1722: The RPC server is unavailable' }
            '1753'  { '1753: There are no more endpoints available from the endpoint mapper' }
            '1818'  { '1818: The remote procedure call was cancelled' }
            '1908'  { '1908: Could not find the domain controller for this domain' }
            '8240'  { '8240: There is no such object on the server' }
            '8333'  { '8333: Directory Object Not Found' }
            '8418'  { '8418: The replication operation failed because of a schema mismatch between the servers involved' }
            '8446'  { '8446: The replication operation failed to allocate memory' }
            '8451'  { '8451: The replication operation encountered a database error' }
            '8453'  { '8453: Replication access was denied' }
            '8464'  { '8464: Synchronization attempt failed' }
            '8477'  { '8477: The replication request has been posted; waiting for reply' }
            '8524'  { '8524: The DSA operation is unable to proceed because of a DNS lookup failure' }
            '8545'  { '8545: Replication update could not be applied' }
            '8589'  { '8589: The DS cannot derive a service principal name (SPN)' }
            '8606'  { '8606: Insufficient attributes were given to create an object' }
            default { $_.LastReplicationResult }
        }
    }

#endregion

#region Get Active Directory Replication Metadata

    $ADReplicationDataParams = @{
        Target = $Domain
        Scope  = 'Domain'
    }

    $ADReplicationMetadata = Get-ADReplicationPartnerMetadata $ADReplicationDataParams |
        Sort-Object -Property @(
            'Server' 
            'Partner'
        )

#endregion

#region Process Active Directory Replication Metadata

    $ADReplicationMetadata = $ADReplicationMetadata |
        Select-Object -Property @(
            @{
                Name       = 'Server'
                Expression = { $_.Server.Split( '.' )[0].ToUpper() }
            }
            @{
                Name       = 'Partner'
                Expression = { $_.Partner.Split( ',' )[1].Split( '=' )[1].ToUpper() }
            }
            @{
                Name       = 'Transport'
                Expression = { $_.IntersiteTransportType } }
            @{
                Name       = 'Last Attempt'
                Expression = { $_.LastReplicationAttempt.ToString( $DateFormat ) }
            }
            @{
                Name       = 'Last Success'
                Expression = { $_.LastReplicationSuccess.ToString( $DateFormat ) }
            }
            @{
                Name       = 'Last Result'
                Expression = { Resolve-StatusCode }
            }
            @{
                Name       = 'Status'
                Expression = { $_.LastReplicationAttempt -match $_.LastReplicationSuccess }
            }
        )

#endregion

#region Prepare Email Body
    
    $HTMLParams = @{
        Fragment   = $true
        PreContent = '<h1>Active Directory Replication Report:</h1>'
    }
    
    $ADReplicationMetadataHTML = $ADReplicationMetadata |
        ConvertTo-Html @HTMLParams
    
    $MailBody = $MailStyle + $ADReplicationMetadataHTML |
        Out-String |
        ForEach-Object -Process {
            $_ `
            -replace '<td>True</td>','<td bgcolor=#99ff99><b>Success</b></td>'`
            -replace '<td>False</td>','<td bgcolor=#ff9999><b>Failure</b></td>'
        }

#endregion

#region Send Email

    If ( $ADReplicationMetadata.LastReplicationResult -notmatch 0 )
    {
        $MailSubject = 'AD Replication Report - Issues Found!'
    }
    Else
    {
        $MailSubject = 'AD Replication Report'
    }
    
    $SendMailParams = @{
        From       = $MailFrom
        To         = $MailTo
        Subject    = $MailSubject
        Body       = $MailBody
        BodyAsHtml = $true
        SmtpServer = $MailServer
    }

    Send-MailMessage @SendMailParams

#endregion

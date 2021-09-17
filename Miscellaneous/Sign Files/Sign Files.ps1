<#
    .AUTHOR
        JC Mocke (https://github.com/jcm-repo/PowerShell/)
    
    .DESCRIPTION
        This script will search for all files with specific file extensions in a directory (including subfolders) and
        display the list in a grid view. A single file or multiple files (hold CTRL while selecting) can then be select
        for signing.

        It is possible to re-sign a file if needed.
    
    .PARAMETERS
        -Path
            Specify a local or network location containing the file/files you would like to sign. Not specifying a 
            location will trigger a directory selection dialog box to open, but this does not support network
            locations.
        
        -Filter
            Specify the file extensions which need to be listed. The defaults are *.ps1, *.psm1, *.psd1, and *.ps1xml.
        
        -SigningAlgorithm
            Supported signing algorithms are MD5, SHA1, SHA256, SHA384, and SHA512. The default is SHA256.
        
        -DisableAudibleBeep
            Disable the audible beep when a message is shown.
        
        -WhatIf
            Files will not be signed.
#>

#region Set Inital Variables

    [CmdletBinding()]
    Param
    (
        [string] $Path,
        
        [array] $Filter = @( 
            '*.ps1'
            '*.psm1'
            '*.psd1'
            '*.ps1xml' ),

        [ValidateSet( 'MD5', 'SHA1', 'SHA256', 'SHA384', 'SHA512', IgnoreCase )]
        [string] $SigningAlgorithm = 'SHA256',

        [switch] $DisableAudibleBeep,

        [switch] $WhatIf
    )

#endregion

#region Set Initial Functions
    
    $Shell = New-Object -ComObject 'WScript.Shell'

    Function Show-Message
    {
        [CmdletBinding()]
        Param
        (
            [string] $Message,

            [ValidateSet( 'Info','Warning','Critical',IgnoreCase )]
            [string] $Type = 'Info',

            [switch] $Quit
        )
        
        If ( !$DisableAudibleBeep.IsPresent )
        {
            [Console]::Beep( 2500, 80 )
        }

        Switch ( $Type )
        {
            'Critical' { $Style = 16 }
            'Warning'  { $Style = 48 }
            'Info'     { $Style = 64 }
        }

        $DialogBox = $Shell.Popup( $Message, 0, $Type, $Style )

        If ( $Quit.IsPresent -or $Type -eq 'Critical' )
        {
            Exit
        }
    }

    Add-Type -AssemblyName System.Windows.Forms

    Function Show-FolderBrowser
    {
        If ( $psISE )
        {
            Write-Output ''
            Write-Warning "PowerShell ISE might cause the folder browser window to open behind other windows!`n"
        }
        
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            RootFolder          = 'MyComputer'
            ShowNewFolderButton = $false
        }

        $DialogResult = $FolderBrowser.ShowDialog()

        If ( $DialogResult -eq 'Cancel' -or !$FolderBrowser.SelectedPath )
        {
            $PathSelectErrorParams = @{
                Message = 'Path was not selected!'
                Type    = 'Critical'
            }

            Show-Message @PathSelectErrorParams
        }
        Else
        {
            $Script:Path = $FolderBrowser.SelectedPath
        }
    }

#endregion

#region Select Certificate
    
    $CertificatesParams = @{
        Path            = 'cert:currentuser\my\'
        CodeSigningCert = $true
    }
    
    $AvailableCertificates = Get-ChildItem @CertificatesParams
        
    If ( !$AvailableCertificates )
    {
        $CertificatesErrorParams = @{
            Message = 'No code signing certificates were found!'
            Type    = 'Critical'
        }
        
        Show-Message @CertificatesErrorParams
    }

    $CertificateGridParams = @{
        Title      = 'Code Signing Certificate Selection'
        OutputMode = 'Single'
    }

    $SelectedCertificate = $AvailableCertificates |
        Out-GridView @CertificateGridParams

    If ( !$SelectedCertificate )
    {
        $CertificateSelectErrorParams = @{
            Message = 'Certificate was not selected!'
            Type    = 'Critical'
        }
        
        Show-Message @CertificateSelectErrorParams
    }

#endregion

#region Select Path

    If ( $Path -and !( Test-Path -Path $Path ) )
    {
        $PathInvalidErrorParams = @{
            Message = 'Specified path is not valid!'
            Type    = 'Warning'
        }

        Show-Message @PathInvalidErrorParams

        Show-FolderBrowser
    }
    ElseIf ( !$Path )
    {
        Show-FolderBrowser
    }

#endregion

#region Build File List

    $GetFilesParams = @{
        Path    = $Path
        Include = $Filter
        Recurse = $true
    }

    $FileList = Get-ChildItem @GetFilesParams |
        Get-AuthenticodeSignature |
        Select-Object -Property @(
            'Status'
            @{
                Name       = 'Issuer'
                Expression = { $_.SignerCertificate.Issuer.Split( ',' )[0].Split( '=' )[1] }
            }
            @{
                Name       = 'File Path'
                Expression = { $_.Path }
            }
        )

    If ( !$FileList )
    {
        $NoFilesFoundErrorParams = @{
            Message = "No files found!`n`n" +
                      "Path: $($Path)`n`n" +
                      "Extension Filter: $( $Filter -join ', ' )"
            Type    = 'Critical'
        }
        
        Show-Message @NoFilesFoundErrorParams
    }

    $FileGridParams = @{
        Title      = "File List - Extention Filter: $( $Filter -join ', ' )"
        OutputMode = 'Multiple'
    }
        
    $SelectedFiles = $FileList |
        Out-GridView @FileGridParams

    If ( !$SelectedFiles )
    {
        $FilesSelectErrorParams = @{
            Message = 'No files were selected!'
            Type    = 'Critical'
        }

        Show-Message @FilesSelectErrorParams
    }

#endregion

#region Sign Files
    
    ForEach ( $File in $SelectedFiles )
    {
        $SignFilesParams = @{
            Certificate   = $SelectedCertificate
            HashAlgorithm = $SigningAlgorithm
            FilePath      = $File.'File Path'
        }
        
        If ( $WhatIf.IsPresent -or $psISE )
        {
            Set-AuthenticodeSignature @SignFilesParams -WhatIf
        }
        Else
        {
            Set-AuthenticodeSignature @SignFilesParams
        }
    }

#endregion

#region Show Results
    
    $ResultsParams = @{
        FilePath = $SelectedFiles.'File Path'
    }
    
    $Results = Get-AuthenticodeSignature @ResultsParams |
        Select-Object -Property @(
            'Status'
            @{
                Name       = 'Issuer'
                Expression = { $_.SignerCertificate.Issuer.Split( ',' )[0].Split( '=' )[1] }
            }
            @{
                Name       = 'File Path'
                Expression = { $_.Path }
            }
        )
    
    $ResultsGridParams = @{
        Title      = 'Signing Results'
        OutputMode = 'None'
    }    
    
    $Results |
        Out-GridView @ResultsGridParams

#endregion

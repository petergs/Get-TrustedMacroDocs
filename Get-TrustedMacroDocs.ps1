﻿function Get-TrustedMacroDocs {
    <#
    .SYNOPSIS
        Gets a list of the macro-enabled documents that have been trusted in Office.
    .DESCRIPTION
        Gets a list of the macro-enabled documents that have been trusted. It accesses the registry
        hives at HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\{OfficeVersion}\{Application}. 
    .PARAMETER Application
        Filter the query to just one of the macro supported Office applications
    .EXAMPLE
        Get-TrustedMacroDocs -Application Excel
        Get all trusted macro-enabled Excel documents
    .LINK
        https://github.com/petergs/Get-TrustedMacroDocs
    #>
    [OutputType([object])]
    param (
        [parameter()][string][ValidateSet('All', 'Excel', 'Word', 'Powerpoint', 'Publisher')] $Application = 'All',
        [parameter()][string][ValidateSet('None', 'Detailed')] $DetailLevel = 'None'
    )

    # Sequence of bytes at the end of the registry value's REG_BINARY Data field that indicates the file has been trusted
    # and macro content is Enabled.
    $trustedByteSeq = @(255, 255, 255, 127)

    $officeVersions = Get-ChildItem -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\" | Select Name -ExpandProperty Name | Where Name -Like "*.0*"

    $apps = @()
    if ($Application -eq 'All') { 
        $apps = ('Excel', 'Word', 'Powerpoint', 'Publisher') 
    }
    else { 
        $apps = ($Application) 
    }

    foreach ($version in $officeVersions) {
        foreach ($app in $apps) {
            $files = $(Get-ItemProperty -Path "Registry::$($version)\$($app)\Security\Trusted Documents\TrustRecords" -ErrorAction SilentlyContinue).PSObject.Properties | Where TypeNameOfValue -eq System.Byte[]
            foreach ($file in $files) { 
                $comp = Compare-Object $file.Value[-4, -3, -2, -1] $trustedByteSeq -PassThru
                if ( $comp -eq $null )  {
                    [PSCustomObject]@{
                        Path          = $file.Name
                        Application   = $app 
                        OfficeVersion = $version.Split('\')[-1]
                    }
                }
            }
        }
    }        
}
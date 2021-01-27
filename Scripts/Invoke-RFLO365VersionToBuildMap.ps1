<#
    .SYSNOPSIS
        Collect and export information about the versions and build of office 365 as well as current status (Expired, Expiring soon or Current)

    .DESCRIPTION
        Collect and export information about the versions and build of office 365 as well as current status (Expired, Expiring soon or Current)

    .PARAMETER GenerateSQLStatement
        Generate a SQL Statement instead of returning an Array

    .NOTES
        Name: Invoke-RFLO365VersionToBuildMap.ps1
        Author: Raphael Perez
        DateCreated: 22 October 2020 (v0.1)

    .EXAMPLE
        $OMList = .\Invoke-RFLO365VersionToBuildMap.ps1 
        .\Invoke-RFLO365VersionToBuildMap.ps1 -GenerateInsertTableSQLStatement | Out-File d:\O365TableVariable.txt
#>
#requires -version 5
[CmdletBinding()]
param(
    [switch]$GenerateDeclaredSQLStatement,
    [switch]$GenerateInsertTableSQLStatement
)

$StartUpVariables = Get-Variable

#region Functions
#region Test-RFLAdministrator
Function Test-RFLAdministrator {
<#
    .SYSNOPSIS
        Check if the current user is member of the Local Administrators Group

    .DESCRIPTION
        Check if the current user is member of the Local Administrators Group

    .NOTES
        Name: Test-RFLAdministrator
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Test-RFLAdministrator
#>
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal $currentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
#endregion

#region Set-RFLLogPath
Function Set-RFLLogPath {
<#
    .SYSNOPSIS
        Configures the full path to the log file depending on whether or not the CCM folder exists.

    .DESCRIPTION
        Configures the full path to the log file depending on whether or not the CCM folder exists.

    .NOTES
        Name: Set-RFLLogPath
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Set-RFLLogPath
#>
    if ([string]::IsNullOrEmpty($script:LogFilePath)) {
        $script:LogFilePath = $env:Temp
    }

    if(Test-RFLAdministrator) {
        # Script is running Administrator privileges
        if(Test-Path -Path 'C:\Windows\CCM\Logs') {
            $script:LogFilePath = 'C:\Windows\CCM\Logs'
        }
    }
    
    $script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
}
#endregion

#region Write-RFLLog
Function Write-RFLLog {
<#
    .SYSNOPSIS
        Write the log file if the global variable is set

    .DESCRIPTION
        Write the log file if the global variable is set

    .PARAMETER Message
        Message to write to the log

    .PARAMETER LogLevel
        Log Level 1=Information, 2=Warning, 3=Error. Default = 1

    .NOTES
        Name: Write-RFLLog
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Write-RFLLog -Message 'This is an information message'

    .EXAMPLE
        Write-RFLLog -Message 'This is a warning message' -LogLevel 2

    .EXAMPLE
        Write-RFLLog -Message 'This is an error message' -LogLevel 3
#>
param (
    [Parameter(Mandatory = $true)]
    [string]$Message,

    [Parameter()]
    [ValidateSet(1, 2, 3)]
    [string]$LogLevel=1)
   
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)) {
        $ScriptName = ''
    } else {
        $ScriptName = $MyInvocation.ScriptName | Split-Path -Leaf
    }

    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($ScriptName):$($MyInvocation.ScriptLineNumber)", $LogLevel
    $Line = $Line -f $LineFormat

    $Line | Out-File -FilePath $script:ScriptLogFilePath -Append -NoClobber -Encoding default
}
#endregion

#region Clear-RFLLog
Function Clear-RFLLog {
<#
    .SYSNOPSIS
        Delete the log file if bigger than maximum size

    .DESCRIPTION
        Delete the log file if bigger than maximum size

    .NOTES
        Name: Clear-RFLLog
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Clear-RFLLog -maxSize 2mb
#>
param (
    [Parameter(Mandatory = $true)][string]$maxSize
)
    try  {
        if(Test-Path -Path $script:ScriptLogFilePath) {
            if ((Get-Item $script:ScriptLogFilePath).length -gt $maxSize) {
                Remove-Item -Path $script:ScriptLogFilePath
                Start-Sleep -Seconds 1
            }
        }
    }
    catch {
        Write-RFLLog -Message "Unable to delete log file." -LogLevel 3
    }    
}
#endregion

#region Get-ScriptDirectory
function Get-ScriptDirectory {
<#
    .SYSNOPSIS
        Get the directory of the script

    .DESCRIPTION
        Get the directory of the script

    .NOTES
        Name: ClearGet-ScriptDirectory
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Get-ScriptDirectory
#>
    Split-Path -Parent $PSCommandPath
}
#endregion

#region Get-RFLOffice365BuildToVersionMap
function Get-RFLOffice365BuildToVersionMap {
<#
    .SYSNOPSIS
        Collect and generate array with information about the versions and build of office 365 as well as current status (Expired, Expiring soon or Current)

    .DESCRIPTION
        Collect and generate array with information about the versions and build of office 365 as well as current status (Expired, Expiring soon or Current)

    .NOTES
        Name: Get-RFLOffice365BuildToVersionMap
        Author: Raphael Perez
        DateCreated: 22 October 2020 (v0.1)
        
        Original Source: https://github.com/ChrisKibble/O365VersionToBuildMap


    .EXAMPLE
        Get-Office365BuildToVersionMap
#>
    # Regular Expression that Finds the Version/Build Numbers from the Page
    [regex]$rxBuilds = '<p><em>Version (.*)<\/em><\/p>'

    # Regular Expression that Finds the Supported version fro the Main page
    [regex]$rxSupport = '<td style="text-align: left;">(.*)<br/></td>'

    # List of all years from 2015 to now.
    # Identify all the possible URLs from 2015 to the Current Year (future proofing?).  Not all of these channels existed in 2015, so a 404 is
    # expected on at least one of them.  There may also not be all pages for the current year if updates haven't been released yet.
    $yearList = 2015..$(get-date).Year

    # Start page for all Office Update pages by Year and Update Type
    $urlBase = "https://docs.microsoft.com/en-us/officeupdates"

    # Main Page to identify supported versions/build
    $MainURL = "$($urlBase)/update-history-microsoft365-apps-by-date"

    # Array to hold the list of Versions and Builds
    $officeBuildList = @()

    #Array to hold the list of the channels with the following format: URL@Channel
    <#Possible URLS: 
    Main Support Page - https://docs.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date
    Current Channel - urls: https://docs.microsoft.com/en-us/officeupdates/current-channel and https://docs.microsoft.com/en-us/officeupdates/current-channel-$year 
    Monthly Enterprise Channel - urls https://docs.microsoft.com/en-us/officeupdates/monthly-enterprise-channel and https://docs.microsoft.com/en-us/officeupdates/monthly-enterprise-channel-$year
    Semi-Annual Enterprise Channel (Preview) - urls https://docs.microsoft.com/en-us/officeupdates/semi-annual-enterprise-channel-preview and https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-targeted-$year
    Semi-Annual Enterprise Channel - urls https://docs.microsoft.com/en-us/officeupdates/semi-annual-enterprise-channel and https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-$year
    Beta channel - url https://docs.microsoft.com/en-us/officeupdates/beta-channel
    Current Channel (Preview) - url https://docs.microsoft.com/en-us/officeupdates/current-channel-preview
    #>
    $urlList = @()
    $urlList += "$($urlBase)/current-channel@Current Channel"
    $urlList += "$($urlBase)/monthly-enterprise-channel@Monthly Enterprise Channel"
    $urlList += "$($urlBase)/semi-annual-enterprise-channel@Semi-Annual Enterprise Channel"
    $urlList += "$($urlBase)/semi-annual-enterprise-channel-preview@Semi-Annual Enterprise Channel (Preview)"
    $urlList += "$($urlBase)/beta-channel@Beta channel"
    $urlList += "$($urlBase)/current-channel-preview@Current Channel (Preview)"

    $yearList | ForEach-Object {
        $urlList += "$($urlBase)/monthly-channel-$($_)@Current Channel"
        $urlList += "$($urlBase)/monthly-enterprise-channel-$($_)@Monthly Enterprise Channel"
        $urlList += "$($urlBase)/semi-annual-channel-$($_)@Semi-Annual Enterprise Channel"
        $urlList += "$($urlBase)/semi-annual-channel-targeted-$($_)@Semi-Annual Enterprise Channel (Preview)"
    }

    #Check the Supported Version  
    Try {
        Write-RFLLog "Connecting to $($MainURL)"
        $web = Invoke-WebRequest -Uri $MainURL -UseBasicParsing
        if($web.StatusCode -ne 200)  {
            Write-RFLLog "$($MainURL) Returned Error $($web.StatusCode)" -LogLevel 2
        } else {
            $content = $web.RawContent
            $content = $content.Substring($content.IndexOf('<p>The following table lists the supported version, and the most current build number, for each update channel.</p>'))
            $content = $content.Substring(0, $content.IndexOf('<p>For information about the approximate download size when updating from'))
            $rxSupportMatches = $rxSupport.matches($content)

            $SupportedVersions = @()
            #$i = Channel
            #$i+1 = Build
            #$i+2 = Version
            #$i+3 = Date Start
            #$i+4 = Date End
            for($i=0; $i -le $rxSupportMatches.Count-5; $i = $i+5) {
                [datetime]$DateStart = $rxSupportMatches[$i+3].Groups[1].Value
                if ([string]($rxSupportMatches[$i+4].Groups[1].Value) -as [DateTime]) {
                    [datetime]$DateEnd = $rxSupportMatches[$i+4].Groups[1].Value
                } else {
                    [datetime]$DateEnd = $DateStart.AddMonths(1)
                }

                if ((($DateEnd - $DateStart).TotalDays) -lt 60) {
                    $Status = 3 #Expire Soon
                } else {
                    $Status = 2 #Current
                }

                $SupportedVersions += New-Object PSObject -Property @{
                    Channel = $rxSupportMatches[$i+0].Groups[1].Value
                    Build = $rxSupportMatches[$i+1].Groups[1].Value
                    Version = $rxSupportMatches[$i+2].Groups[1].Value
                    Status = $Status
                    DateStart = $DateStart
                    DateEnd = $DateEnd
                }
            }
        }
    } catch {
        Write-RFLLog "Unknown error occurred: $($_)" -LogLevel 2
    }

    ForEach($item in $urlList) {
        $itemArray = $item.Split('@')
        $url = $itemArray[0]
        $channel = $itemArray[1]

        Write-RFLLog "Connecting to channel $($Channel) url: $($url)"
        Try {
            $web = Invoke-WebRequest -Uri $url -UseBasicParsing
        } catch {
            <# #>
        }

        if($web.StatusCode -ne 200)  {
            Write-RFLLog "$($url) Returned Error $($web.StatusCode)" -LogLevel 2
        } else {
            $content = $web.RawContent
            $rxMatches = $rxBuilds.matches($content)

            ForEach($entry in $rxMatches) {
                $buildLine = $entry.Groups[1].Value
                $buildNumber = $buildLine.substring(0,$buildLine.indexOf(' '))
                $versionFromWeb = $($buildLine.substring($buildLine.indexOf('Build') + 6)) -replace '\)',''

                $IsSupport = $SupportedVersions | Where-Object {($_.Version -like "$($versionFromWeb.split('.')[0]).*") -and ($_.Channel -eq $channel) -and ($_.Build -eq $buildNumber)}
                if ($IsSupport) {
                    $Status = $IsSupport.Status
                } else {
                    $Status = 4
                }
                [version]$versionNumber = "16.0.$versionFromWeb"

                Write-RFLLog "Adding BuildNumber $($buildNumber), VersionNumber $($versionNumber), Channel $($channel), Status ($Status)"
                $officeBuildList += New-Object PSObject -Property @{
                    BuildNumber = $buildNumber
                    VersionNumber = $versionNumber
                    Channel = $channel
                    Status = $Status
                }

                #Microsoft has renamed some channels, so making sure old names match. 
                #I'm only adding the ones i've found on my environment
                #use the below select to check what you have againt the ConfigMgr database
                #select distinct Channel0 from fn_rbac_GS_OFFICE_PRODUCTINFO('disabled') 
                #https://docs.microsoft.com/en-us/DeployOffice/update-channels-changes
                if ($channel -eq 'Current Channel') {
                    $officeBuildList += New-Object PSObject -Property @{
                        BuildNumber = $buildNumber
                        VersionNumber = $versionNumber
                        Channel = 'Monthly'
                        Status = $Status
                    }
                }
            }
        }
    }

    $officeBuildList
}
#endregion
#endregion

#region Variables
$script:ScriptVersion = '0.1'
$script:LogFilePath = $env:Temp
$Script:LogFileFileName = 'Invoke-RFLO365VersionToBuildMap.log'
$script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
$Script:tableName = "O365BuildToVersionMap"
#endregion

#region Main
try {
    Set-RFLLogPath
    Clear-RFLLog 25mb

    Write-RFLLog -Message "*** Starting ***"
    Write-RFLLog -Message "Script version $($script:ScriptVersion)"
    Write-RFLLog -Message "Running as $($env:username) $(if(Test-RFLAdministrator) {"[Administrator]"} Else {"[Not Administrator]"}) on $($env:computername)"

    $PSCmdlet.MyInvocation.BoundParameters.Keys | ForEach-Object { 
        Write-RFLLog -Message "Parameter '$($_)' is '$($PSCmdlet.MyInvocation.BoundParameters.Item($_))'"
    }

    $OfficeList = Get-RFLOffice365BuildToVersionMap

    $sql = @()
    if ($GenerateDeclaredSQLStatement) {
        # Declare the table
        $sql += "DECLARE @$($Script:tableName) TABLE (VersionNumber VARCHAR(32) NOT NULL, BuildNumber VARCHAR(8) NOT NULL, Channel VARCHAR(50) NOT NULL, Status INT NOT NULL);"

        $OfficeList | ForEach-Object {
            $sql += "INSERT INTO @$($Script:tableName) VALUES ('$($_.VersionNumber)','$($_.BuildNumber)','$($_.Channel)',$($_.Status));"
        }

        $sql
    } elseif ($GenerateInsertTableSQLStatement) {
        # Declare the table
        $sql += "IF OBJECT_ID('$Script:tableName', 'U') IS NOT NULL DROP TABLE $Script:tableName;"

        $sql += "CREATE TABLE $Script:tableName (VersionNumber VARCHAR(32) NOT NULL, BuildNumber VARCHAR(8) NOT NULL, Channel VARCHAR(50) NOT NULL, Status INT NOT NULL);"

        $OfficeList | ForEach-Object {
            $sql += "INSERT INTO $($Script:tableName) VALUES ('$($_.VersionNumber)','$($_.BuildNumber)','$($_.Channel)',$($_.Status));"
        }

        $sql
    } else {
        $OfficeList
    }
} catch {
    Write-RFLLog -Message "An error occurred $($_)" -LogLevel 3
    return 3000
} finally {
    Get-Variable | Where-Object { $StartUpVariables.Name -notcontains $_.Name } |
    ForEach-Object {
        Try { 
            Write-RFLLog -Message "Removing Variable $($_.Name)"
            Remove-Variable -Name "$($_.Name)" -Force -Scope "global" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        } Catch { 
            Write-RFLLog -Message "Unable to remove variable $($_.Name)"
        }
    }
    Write-RFLLog -Message "*** Ending ***"
}
#endregion

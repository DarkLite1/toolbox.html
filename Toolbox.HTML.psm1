#Requires -Modules ActiveDirectory

$SMTPserver = $env:SMTP_SERVER
$Quotes = 'T:\Input\Extra\Quotes.txt'

try {
    if (-not $SMTPserver) {
        throw 'SMTP server name required'
    }
}
catch {
    throw "Failed loading the module 'Toolbox.HTML': $_"
}

Function ConvertTo-HTMLlinkHC {
    <#
    .SYNOPSIS
        Converts a path or URL to an HTML <a href> tag.

    .DESCRIPTION
        Converts a path or URL to a clickable link and generates the correct HTML code: '<a href="LINK/PATH/URL">Click me</a>'.

    .PARAMETER Name
        The name that represents the link, it's the text that people see and 
        on what they click to open the link  '<a href="LINK/PATH/URL">Name</a>'.
        
        By default the 'Name' will be formatted so the first letter is in 
        upper case and the rest in lower case. In case the name does not need 
        formatting use 'FormatName = $false'.

        -Name 'gOoGLe'      > 'Google'
        -Name 'self help'   > 'Self help'

    .PARAMETER Path
        Location of a folder, file or URL.

    .PARAMETER FormatName
        The field 'Name' is always formatted so the first is in upper case and 
        the rest in lower case. 
        
        Valid options:
        -Name 'lOg FoLdEr'                     > 'Log folder' (default)
        -Name 'lOg FoLdEr' -FormatName $False  > 'lOg FoLdEr'

    .EXAMPLE
        $params = @{
            Name = 'scheduled Task'
            Path = '\\PC1\Log\SchedUleD TaSk'    
        }
        ConvertTo-HTMLlinkHC @params

        Create string '<a href="\\PC1\log\scheduled task">Scheduled task</a>'

    .EXAMPLE
        $params = @{
            Name = 'google'
            Path = 'http://www.google.com'
        }
        ConvertTo-HTMLlinkHC @params

        Create string '<a href="http://www.google.com">Google</a>'

    .EXAMPLE
        $params = @{
            Name       = 'This is IMPORTANT!'
            Path       = 'T:\Reports\Fruits\Best fruits of 2014.xlsx'
            FormatName = $false
        }
        ConvertTo-HTMLlinkHC @params

        Create string '<a href="T:\Reports\fruits\best fruits of 2014.xlsx">This is IMPORTANT!</a>'
    #>

    [CmdletBinding()]
    Param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [String]$Name,
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [Alias("Link", "URL")]
        [String]$Path,
        [parameter(Mandatory = $false)]
        [ValidateSet($True, $False)]
        [String]$FormatName = $True
    )

    Process {

        Switch -Regex ($Path) {
            '^[\\]' { $Path = $Path.Substring(0, 2) + $Path.Substring(2, 1).ToUpper() + $Path.Substring(3).ToLower() }  # UNC-Path
            '^[a-z]:' { $Path = $Path.Substring(0, 1).ToUpper() + $Path.Substring(1, 2) + $Path.Substring(3, 1).ToUpper() + $Path.Substring(4).ToLower() } # Local path
        }

        if ($FormatName -eq $True) {
            Switch -Regex ($Name) {
                '^[\\]' { $Name = $Name.Substring(0, 2) + $Name.Substring(2, 1).ToUpper() + $Name.Substring(3).ToLower() }  # UNC-Path
                '^[a-z]:' { $Name = $Name.Substring(0, 1).ToUpper() + $Name.Substring(1, 2) + $Name.Substring(3, 1).ToUpper() + $Name.Substring(4).ToLower() } # Local path
                Default { $Name = $Name.Substring(0, 1).ToUpper() + $Name.Substring(1).ToLower() }
            }
        }

        $HTMLstring = @"
<a href="$Path">$Name</a>
"@
        Write-Output $HTMLstring
    }
}
Function ConvertTo-HtmlListHC {
    <#
    .SYNOPSIS
        Creates an unordered HTML list.

    .DESCRIPTION
        Creates an unordered HTML list from an array to use in an HTML 
        document, like an e-mail or a HTML file.

    .PARAMETER Message
        The items in the list.

    .PARAMETER Spacing
        Defines how close the items are together. T
        
        Valid options:
        'Normal': Close together, no breaks in between (default)
        'Wide'  : Double breaks between each item

    .PARAMETER Header
        Add a header '<h3>My list title</h3>' above the unordered list.

    .PARAMETER FootNote
        Add a small text at the bottom of the unordered list in a smaller font 
        and italic. This is convenient for adding a small explanation of the 
        items or a legend.

    .EXAMPLE
        $params = @{
            Message = @('Item 3', 'Item 1', 'Item 2')
        }
        ConvertTo-HTMLlistHC @params
        
        Create the following HTML code:
        '<ul>
            <li>Item 1</li>
            <li>Item 2</li>
            <li>Item 3</li>
        </ul>'

    .EXAMPLE
        $params = @{
            Message  = @('Apples', 'Peers', 'Bananas')
            Spacing  = 'Wide'
            FootNote = 'These are all tasting sweet'
        }
        ConvertTo-HTMLlistHC @params
        
        Create the following HTML code:
        '<ul>
            <li>Apples<br><br></li>
            <li>Bananas<br><br></li>
            <li>Peers<br><br>
            <i><font size="2">* These are all tasting sweet</font></i></li>
        </ul>'
    #>

    [CmdletBinding()]
    Param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Message,
        [ValidateSet('Normal', 'Wide')]
        [String]$Spacing = 'Normal',
        [String]$Header,
        [String]$FootNote,
        [Switch]$NoSorting
    )

    Begin {
        $HTMLlist = New-Object System.Collections.ArrayList
        $Space = Switch ($Spacing) {
            'Normal' {}
            'Wide' { '<br><br>' }
        }
        if ($FootNote) {
            $Footer = "<br><br><i><font size=`"2`">* $FootNote</font></i>"
        }
    }

    Process {
        $Message | ForEach-Object {
            $null = $HTMLlist.Add('<li>' + $_ + "$Space</li>")
        }
    }

    End {
        if ($NoSorting) {
            $HTMLlist = , "<ul>$HTMLlist</ul>"
        }
        else {
            $HTMLlist = , "<ul>$($HTMLlist | Sort-Object)</ul>"
        }

        $HTMLlist = $HTMLlist.Replace("$Space</li></ul>", "$Footer</li></ul>") # Remove last double breaks

        if ($Header) {
            $HTMLlist = , "<h3>$Header</h3>$HTMLlist"
        }

        Write-Output $HTMLlist
    }
}
Function Get-ScriptRuntimeHC {
    <#
    .SYNOPSIS
        Calculates the total runtime of the script.

    .DESCRIPTION
        This function checks how long the script has been running by using the 
        switches 'Start' and 'Stop' in the beginning and end of the script.

    .PARAMETER Start
        Start time when the script starts. Usually put in the beginning of the 
        script.

    .PARAMETER Stop
        Stop time when the script stops. Usually put at the end of the script.

    .EXAMPLE
        $null = Get-ScriptRuntime -Start

        Start-Sleep -Seconds 5

        $null = Get-ScriptRuntime -Stop
        Send-MailHC -Message ....

        Saves the script start time in a global variable, at the end of the script execution the elapsed time is calculated an used in the footer of
        the e-mail send.
    #>

    [CmdletBinding()]
    Param (
        [parameter(Mandatory = $true, ParameterSetName = "StartSwitch")]
        [Switch]$Start,
        [parameter(Mandatory = $true, ParameterSetName = "StopSwitch")]
        [Switch]$Stop
    )

    if ($Start) {
        $Script:ScriptStartTime = (Get-Date)
        Write-Verbose "Get-ScriptRunTime: Start time`t`t: $Script:ScriptStartTime"
        Write-Output $Script:ScriptStartTime
    }

    if ($Stop) {
        $Script:ScriptEndTime = (Get-Date)
        Write-Verbose "Get-ScriptRunTime: Stop time`t`t: $Script:ScriptEndTime"
        $RunTime = New-TimeSpan -Start $Script:ScriptStartTime -End $Script:ScriptEndTime
        $Script:ScriptRunTime = ("{0:00}:{1:00}:{2:00}" -f $RunTime.Hours, $RunTime.Minutes, $RunTime.Seconds)
        Write-Verbose "Get-ScriptRunTime: Total run time`t: $Script:ScriptRunTime"
        Write-Output $Script:ScriptRunTime
    }
}
Function New-LogFileNameHC {
    <#
    .SYNOPSIS
        Generates strings that can be used as a file name.

    .DESCRIPTION
        Converts strings or paths to usable formats for file names and adds the 
        date if required. It filters out all the unaccepted characters by 
        Windows to use a UNC-path or local-path as a file name. It's also 
        useful for adding the date to a string. In case a path is provided, the 
        first letter will be in upper case and the rest will be in lower case. 
        It will also check if the log file already exists, and if so, create a 
        new one with an increased number [0], [1], ..

    .PARAMETER LogFolder
        Folder path where the log files are located.

    .PARAMETER Name
        Can be a path name or just a string.

    .PARAMETER Date
        Adds the date to the name. When using one of the 'Script-options', make 
        sure to use 'Get-ScriptRuntime (Start/Stop)' in your script. 
        
        Valid options:
        'ScriptStartTime' : Start time of the script
        'ScriptEndTime'   : End time of the script
        'CurrentTime'     : Time when the command ran

    .PARAMETER Location
        Places the selected date in front or at the end of the name.
        
        Valid options:
        - Begin : 2014-09-25 - Name (default)
        - End   : Name - 2014-09-25

    .PARAMETER Format
        Format used for the selected date. 
        
        Valid options:
        - yyyy-MM-dd HHmm (DayOfWeek)   : 2014-09-25 1431 (Thursday) (default)
        - yyyy-MM-dd HHmmss (DayOfWeek) : 2014-09-25 143121 (Thursday)
        - yyyyMMdd HHmm (DayOfWeek)     : 20140925 1431 (Thursday)
        - yyyy-MM-dd HHmm               : 2014-09-25 1431
        - yyyyMMdd HHmm                 : 20140925 1431
        - yyyy-MM-dd                    : 2014-09-25
        - yyyyMMdd                      : 20140925

    .PARAMETER NoFormatting
        Doesn't change the string to phrase format with a capital in the 
        beginning. However, it still removes/replaces all characters that are 
        not allowed in a Windows file name.

    .PARAMETER Unique
        When this switch is set, we will first check if a file exists with the 
        same name. If it does, we add a number to the file, every time it runs 
        the counter will go up.

    .EXAMPLE
        $params = @{
            LogFolder = 'T:\Log folder'
            Name      = 'Drivers.log'
            Date      = 'CurrentTime'
            Position  = 'End'
        }
        New-LogFileNameHC @params

        Create the string 'T:\Log folder\Drivers - 2015-01-26 1028 (Monday).log'

    .EXAMPLE
        $params = @{
            LogFolder = 'T:\Log folder'
            Format    = 'yyyyMMdd'
            Name      = 'Drivers.log'
            Date      = 'CurrentTime'
            Position  = 'Begin'
        }
        New-LogFileNameHC @params

        Create the string 'T:\Log folder\20220621 - Drivers.log'

    .EXAMPLE
        $params = @{
            LogFolder = 'T:\Log folder'
            Name      = 'Drivers.log'
        }
        New-LogFileNameHC @params

        Create the string 'T:\Log folder\Drivers.log'
    #>

    [CmdletBinding()]
    Param (
        [parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Set1')]
        [parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Set2')]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [String]$LogFolder,
        [parameter(Mandatory = $true, Position = 1, ParameterSetName = 'Set1', ValueFromPipeline = $true)]
        [parameter(Mandatory = $true, Position = 1, ParameterSetName = 'Set2', ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [alias('Path')]
        [String[]]$Name,
        [parameter(Mandatory = $true, Position = 2, ParameterSetName = 'Set2')]
        [ValidateSet('ScriptStartTime', 'ScriptEndTime', 'CurrentTime')]
        [String]$Date,
        [parameter(Mandatory = $false, Position = 3, ParameterSetName = 'Set2')]
        [ValidateSet('Begin', 'End')]
        [alias('Location')]
        [String]$Position = 'Begin',
        [parameter(Mandatory = $false, Position = 4, ParameterSetName = 'Set2')]
        [ValidateSet('yyyy-MM-dd HHmm (DayOfWeek)', 'yyyy-MM-dd HHmmss (DayOfWeek)',
            'yyyyMMdd HHmm (DayOfWeek)', 'yyyy-MM-dd HHmm', 'yyyyMMdd HHmm',
            'yyyy-MM-dd', 'yyyyMMdd')]
        [String]$Format = 'yyyy-MM-dd HHmm (DayOfWeek)',
        [Switch]$NoFormatting,
        [Switch]$Unique
    )

    Begin {
        if ($Date) {
            Switch ($Date) {
                'ScriptStartTime' { $d = $ScriptStartTime }
                'ScriptEndTime' { $d = $ScriptEndTime }
                'CurrentTime' { $d = Get-Date }
            }

            Switch ($Format) {
                'yyyy-MM-dd HHmm (DayOfWeek)' {
                    $DateFormat = "{0:00}-{1:00}-{2:00} {3:00}{4:00} ({5})" `
                        -f $d.Year, $d.Month, $d.Day, $d.Hour, $d.Minute, $d.DayOfWeek
                }
                'yyyy-MM-dd HHmmss (DayOfWeek)' {
                    $DateFormat = "{0:00}-{1:00}-{2:00} {3:00}{4:00}{5:00} ({6})" `
                        -f $d.Year, $d.Month, $d.Day, $d.Hour, $d.Minute, $d.Second, $d.DayOfWeek
                }
                'yyyyMMdd HHmm (DayOfWeek)' {
                    $DateFormat = "{0:00}{1:00}{2:00} {3:00}{4:00} ({5})" `
                        -f $d.Year, $d.Month, $d.Day, $d.Hour, $d.Minute, $d.DayOfWeek
                }
                'yyyy-MM-dd HHmm' { $DateFormat = ($d).ToString("yyyy-MM-dd HHmm") }
                'yyyyMMdd HHmm' { $DateFormat = ($d).ToString("yyyyMMdd HHmm") }
                'yyyy-MM-dd' { $DateFormat = ($d).ToString("yyyy-MM-dd") }
                'yyyyMMdd' { $DateFormat = ($d).ToString("yyyyMMdd") }
            }

            Switch ($Position) {
                'Begin' { $Prefix = "$DateFormat - " }
                'End' { $Postfix = " - $DateFormat" }
            }
        }
    }

    Process {
        foreach ($N in $Name) {
            if ($N -match '[.]...$|[.]....$') {
                $Extension = ".$($N.Split('.')[-1])"
                $N = $N.Replace("$Extension", '')
            }

            if ($N -match '[\\]') {
                $Path = $N -replace '\\', '_'
                $Path = $Path -replace ':', ''
                $Path = $Path -replace ' ', ''
                $Path = $Path.TrimStart("__")

                if ($NoFormatting) {
                    $N = $Path
                }
                else {
                    if ($Path -match '^[a-z]_') {
                        $N = $Path.Substring(0, 1).ToUpper() + $Path.Substring(1, 1) +
                        $Path.Substring(2, 1).ToUpper() + $Path.Substring(3).ToLower() # Local path
                    }
                    else {
                        $N = $Path.Substring(0, 1).ToUpper() + $Path.Substring(1).ToLower() # UNC-path
                    }
                }
            }
            else {
                if ($NoFormatting) {
                    $N = $N
                }
                else {
                    $N = $N.Substring(0, 1).ToUpper() + $N.Substring(1).ToLower()
                }
            }

            if ($Unique) {
                $FileName = "$LogFolder\$Prefix$N$Postfix{0}$Extension"

                Function Increment-Index ($f) {
                    $parts = $f.Split('{}')
                    "$($parts[0]){$((1 + $parts[1]))}$($parts[2])"
                }

                while (Test-Path -LiteralPath $FileName) {
                    $FileName = Increment-Index $FileName
                }
            }
            else {
                $FileName = "$LogFolder\$Prefix$N$Postfix$Extension"
            }

            Write-Output $FileName
        }
    }
}
Function Send-MailHC {
    <#
    .SYNOPSIS
        Send an e-mail message as anonymous, when allowed by the SMTP-Relay 
        server.

    .DESCRIPTION
        This function sends out a preformatted HTML e-mail by only providing 
        the recipient, subject and body. The e-mail formatting is optimized for 
        MS Outlook. All e-mails sent will be stored in the Windows Event Log 
        under the event log of the ScriptName (Header parameter).

    .PARAMETER From
        The sender address, by preference this is an existing mailbox so we get 
        the bounce back mail in case of failure.

    .PARAMETER To
        The e-mail address of the recipient(s) you wish to e-mail.

    .PARAMETER Bcc
        The e-mail address of the recipient(s) you wish to e-mail in Blind 
        Carbon Copy. Other users will not see the e-mail address of users in 
        the 'Bcc' field.

    .PARAMETER Cc
        The e-mail address of the recipient(s) you wish to e-mail in Carbon 
        Copy.

    .PARAMETER From
        The e-mail address from which the mail is sent. If not specified, 
        the default value will be the script name or 'Test' when the script 
        name is unknown.

    .PARAMETER Subject
        The Subject-header used in the e-mail.

    .PARAMETER Message
        The message you want to send will appear by default in a paragraph 
        '<p>My message</p>'. If you want to have a title/header to, you can use:
        -Message "<h3>Header one:<\h3>", "My message"

    .PARAMETER Priority
        Specifies the priority of the e-mail message. 
        Valid values
        - Normal (default)
        - High
        - Low

    .PARAMETER Attachments
        Specifies the path and file names of files to be attached to the e-mail 
        message.

    .PARAMETER LogFolder
        Specifies the location where the log files are stored.

    .PARAMETER PSEmailServer
        The SMTP server used to send mails.

    .PARAMETER Save
        The full path file where the e-mail will be saved in HTML format.

    .EXAMPLE
        $params = @{
            To      = 'Bob@domain.com'
            Subject = 'Flavor report'
        }
        @(
            '<h3>Fruits:</h3>', 
            'Peers and apples are great',
            '<h3>Vegetables:</h3>',
            'Aubergines not so much'
        ) | Send-MailHC @params

        Bob will receive an e-mail with the subject 'Flavor report' and the a 
        summary of fruits and vegetables with each their own header/title.

    .EXAMPLE
        $params = @{
            To       = @('Bob@domain.com', 'Jack@Reacher.com')
            Subject  = 'Summary'
            Message  = 'You did great today. Thank you for participating' 
            Priority = 'High'
        }
        Send-MailHC @params

        Bob and Jack receive an e-mail with high priority to inform them that 
        they did a good job.
    #>

    [CmdLetBinding()]
    Param (
        [parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String[]]$To,
        [parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]$Subject,
        [parameter(Mandatory, Position = 2, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Message,
        [String[]]$Cc,
        [String[]]$Bcc,
        [String]$Header = 'Test',
        [ValidateScript( { Test-Path $_ -PathType Container })]
        [IO.DirectoryInfo]$LogFolder,
        [ValidateScript( { Test-Path $_ -PathType Leaf })]
        [String[]]$Attachments,
        [ValidateSet('Low', 'Normal', 'High')]
        [String]$Priority = 'Normal',
        [String]$SMTPserver = $SMTPserver,
        [ValidateNotNullOrEmpty()]
        [String]$From = "PowerShell@$env:COMPUTERNAME",
        [IO.FileInfo]$Save,
        [String]$EventLogSource = $ScriptName,
        [String]$EventLogName = 'HCScripts',
        [String]$Quotes = $Quotes
    )

    Begin {
        Function Test-MailExistInDomainHC {
            <#
            .SYNOPSIS
                Check if a mail address exists.
        
            .DESCRIPTION
                Check if a mail address exists within the active directory.
        
            .EXAMPLE
                Test-MailExistInDomainHC Alerts@sagrex.be
                Returns data when it exists and nothing when it doesn't exist.
            #>
        
            Param (
                [Parameter(Mandatory)]
                [String]$Mailbox
            )
        
            $Filter = "mail -eq ""$Mailbox"" -or proxyAddresses -eq ""smtp:$Mailbox"""
            Get-ADObject -Properties mail, proxyAddresses -Filter $Filter
        }

        Try {
            $EncUTF8 = New-Object System.Text.utf8encoding

            $OriginalMessage = @()

            #region Check From address to make sure mails arrive
            if (Test-MailExistInDomainHC $From) {
                Write-Warning "Send-MailHC: The header '$Header' will not be visible in MS Outlook, the DisplayName of '$From' will be visible instead because the sender account is known in the GAL."
            }
            else {
                foreach ($T in $To) {
                    if (-not (Test-MailExistInDomainHC $T)) {
                        throw "Mail address '$T' not found in AD and sending mail to external mail addresses is not supported when the 'From' address is not found in AD."
                    }
                }
            }
            #endregion

            #region Excel files that are opened can't be sent as attachment, so we copy them first
            $Attachment = New-Object System.Collections.ArrayList($null)

            $TmpFolder = "$env:TEMP\Send-MailHC {0}" -f (Get-Random)
            foreach ($a in $Attachments) {
                if ($a -like '*.xlsx') {
                    if (-not(Test-Path $TmpFolder)) {
                        $null = New-Item $TmpFolder -ItemType Directory
                    }
                    Copy-Item $a -Destination $TmpFolder

                    $null = $Attachment.Add("$TmpFolder\$(Split-Path $a -Leaf)")
                }
                else {
                    $null = $Attachment.Add($a)
                }
            }
            #endregion
        }
        Catch {
            $Global:Error.RemoveAt(0)
            throw "Failed sending e-mail to '$To': $_"
        }
    }

    Process {
        Foreach ($M in $Message) {
            $M = $M.Trim()

            $OriginalMessage += $M
            if ($M -like '<*') {
                # We receive pre-formatted HTML-code
                $Messages += $M
            }
            else {
                # We assume normal text is being sent and put it in a paragraph
                $Messages += "<p>$M</p>"
            }
        }
    }

    End {
        Try {
            $HTML = @"
<!DOCTYPE html>
<html><head><style type="text/css">
body {font-family:verdana;background-color:white;}
h1 {background-color:black;color:white;margin-bottom:10px;text-indent:10px;page-break-before: always;}
h2 {background-color:lightGrey;margin-bottom:10px;text-indent:10px;page-break-before: always;}
h3 {background-color:lightGrey;margin-bottom:10px;font-size:16px;text-indent:10px;page-break-before: always;}
p {font-size: 14px;margin-left:10px;}
p.italic {font-style: italic;font-size: 12px;}
table {font-size:14px;border-collapse:collapse;border:1px none;padding:3px;text-align:left;padding-right:10px;margin-left:10px;}
td, th {font-size:14px;border-collapse:collapse;border:1px none;padding:3px;text-align:left;padding-right:10px}
li {font-size: 14px;}
base {target="_blank"}
</style></head><body>
<h1>$Header</h1>
<h2>The following has been reported:</h2>
$Messages
<h2>About</h2>
<table>
<colgroup><col/><col/></colgroup>
$(if($ScriptStartTime){$("<tr><th>Start time</th><td>{0:00}/{1:00}/{2:00} {3:00}:{4:00} ({5})</td></tr>" -f `
$ScriptStartTime.Day,$ScriptStartTime.Month,$ScriptStartTime.Year,$ScriptStartTime.Hour,$ScriptStartTime.Minute,$ScriptStartTime.DayOfWeek)})
$(if($ScriptRunTime){"<tr><th>Total runtime</th><td>$ScriptRunTime $(if($MaxThreads){"($MaxThreads jobs at once)"})</td></tr>"})
$(
    if ($LogFolder) {
        "<tr><th>Log folder</th><td>$("<a href=`""$LogFolder"`">Open log folder</a>")</td></tr>"
    }
)
$(
    if ($ImportFile) {
        "<tr><th>Import file</th><td>$("<a href=`""$ImportFile"`">$ImportFile</a>")</td></tr>"
    }
)
$(if($global:PSCommandPath){"<tr><th>PSCommandPath</th><td>$global:PSCommandPath</td></tr>"})
<tr><th>Host</th><td>$($host.Name)</td></tr>
<tr><th>ComputerName</th><td>$env:COMPUTERNAME</td></tr>
<tr><th>Whoami</th><td>$("$env:USERDNSDOMAIN\$env:USERNAME")</td></tr>
</table>
$(
    if (($Quotes) -and (Test-Path $Quotes)) {
        '<p class=italic>"' + $(Get-Content $Quotes | Get-Random -ErrorAction SilentlyContinue) + '"</p>'
    }
)
</body></html>
"@

            $EmailParams = @{
                To          = $To
                Cc          = $Cc
                Bcc         = $Bcc
                From        = $Header + ' <' + $From + '>'
                Subject     = $Subject
                Body        = $HTML
                BodyAsHtml  = $True
                Priority    = $Priority
                SMTPServer  = $SMTPserver
                Attachments = $Attachment
                Encoding    = $EncUTF8
                ErrorAction = 'Stop'
            }

            #region Remove empty params
            $list = New-Object System.Collections.ArrayList($null)

            foreach ($h in $EmailParams.Keys) { 
                if ($($EmailParams.Item($h)) -eq $null) {
                    $null = $list.Add($h) 
                } 
            }
            foreach ($h in $list) {
                $EmailParams.Remove($h)
            }
            #endregion

            Send-MailMessage @EmailParams
            Write-Verbose "Mail sent to '$To'"
        }
        Catch {
            $Global:Error.RemoveAt(0)
            throw "Failed sending e-mail to '$($To)': $_"
        }
        Finally {
            if (Test-Path $TmpFolder) {
                Remove-Item -LiteralPath $TmpFolder -Recurse -Force
            }
        }

        Try {
            #region Save in event log

            # Limit the message text we capture in the Windows Event Log
            $Text = $OriginalMessage | Out-String

            $TextCharsToSave = 600

            if (($End = $Text.Length) -gt $TextCharsToSave) {
                $End = $TextCharsToSave
            }

            if (-not $EventLogSource) {
                $EventLogSource = 'test'
            }

            if (
                -not(
                    ([System.Diagnostics.EventLog]::Exists($EventLogName)) -and
                    [System.Diagnostics.EventLog]::SourceExists($EventLogSource)
                )
            ) {
                $newEventLogParams = @{
                    LogName     = $EventLogName
                    Source      = $EventLogSource
                    ErrorAction = 'Stop'
                }
                New-EventLog @newEventLogParams
            }
    
            $eventLogParams = @{
                LogName     = $EventLogName
                Source      = $EventLogSource
                EntryType   = 'Information'
                EventID     = '1' 
                ErrorAction = 'Stop'
            }

            Write-EventLog @eventLogParams -Message (
                'Mail sent' + "`n`n" +
                "- Subject:`t" + $EmailParams.Subject + "`n" +
                "- To:`t`t" + $EmailParams.To + "`n" +
                "`n" + $Text.Substring(0, $End) + "`n`n" +
                "- CC:`t`t" + $EmailParams.Cc + "`n" +
                "- BCC:`t`t" + $EmailParams.Bcc + "`n" +
                "- Priority:`t" + $EmailParams.Priority + "`n" +
                "- SMTPServer:`t" + $EmailParams.SMTPServer + "`n" +
                "- From:`t`t" + $EmailParams.From + "`n" +
                "- Attachments:`t" + $EmailParams.Attachments + "`n" +
                "- Start time:`t" + $(if ($ScriptStartTime) { $ScriptStartTime.ToString('dd/MM/yyyy HH:mm:ss (dddd)') }) + "`n" +
                "- Total runtime:`t" + $ScriptRunTime + "`n" +
                "- LogFolder:`t" + $LogFolder + "`n" +
                "- Import file:`t" + $ImportFile + "`n" +
                "- Script location:`t" + $global:PSCommandPath
            )
            #endregion
        }
        Catch {
            $Global:Error.RemoveAt(0)
            throw "E-mail sent successfully but couldn't save the event in the Windows Event Log: $_"
        }

        if ($Save) {
            Try {
                Out-File -FilePath $Save -InputObject $EmailParams.Body -Encoding utf8
                Write-Verbose "Mail saved in '$Save'"
            }
            Catch {
                $Global:Error.RemoveAt(0)
                throw "E-mail sent successfully but couldn't save the e-mail with full path name '$Save': $_"
            }
        }
    }
}

Export-ModuleMember -Function * -Alias *
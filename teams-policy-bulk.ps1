<#
.SYNOPSIS
Designed to do bulk application of Teams calling and messaging policies to 0365 users.

.DESCRIPTION
 REQUIREMENTS:	Use -requirements
 USAGE: 		Use -help or run get-help .\script.ps1
				
 AUTHOR:		James Tuson
 DATE CREATED: 	Use -changelog 
 LAST UPDATED: 	Use -changelog

.PARAMETER ResumeAt
The row number at which to resume processing the target CSV file.

.PARAMETER PathToCSV
Full path name to the csv input file. Needs headers. Required headers are Mail

.INPUTS
csv file containing list of UPNs or email addresses to which the policy/ies should be applied.

.OUTPUTS

.EXAMPLE
PS> .\teams-policy-bulk.ps1 -resumeat 1123 -csv "\\domain\dfs\share\all_students_UPN.csv
PS> get-help .\teams-policy-bulk.ps1 -full

.NOTES
	It is advised to manually test running the Grant-CsTeams...commands
	first using an example target user to ensure the credentials work, the 
	policy names are correct and the session is established correctly.
	Note that the o365 session will timeout after 1hr. For me that was at around
	row 1500. Hence I've included the -resumeat parameter. I had to close then
	reopen the powershell window before I could re-establish the session.
	
.LINK
https://github.com/gostega/teams-bulk-policy

#note - there need to be at least two empty lines below the end of the help
#>


param (
	
	#script agnostic paramaters
	[Alias('test')][switch]$param_test, 
	[Alias('help','?')][switch]$param_help, 
	[Alias('features','info')][switch]$param_features,
	[Alias('requirements')][switch]$param_requirements,
	[Alias('knownissues','bugs','issues')][switch]$param_issues,
	[Alias('changelog','changes')][switch]$param_changelog,
	[Alias('logfile','log','logpath','logtofile')][System.IO.FileInfo]$param_logfilepath,
	
	#script specific parameters
	[Alias('resumeat','resumeindex')]
	[int]
	$param_resumeat=1
	,
	
	[Alias('csv','pathtocsv','importpath')]
	[System.IO.FileInfo]
	$param_csvpath
	,
	
	[Alias('group','adgroup','usergroup')]
	[string]
	$param_userADgroup
	,
	
	[Alias('ou','userOU')]
	[string]
	$param_userOU
	,
	
	[Alias('messagingpolicy','messaging')]
	[string]
	$param_messagingpolicy
	,
	
	[Alias('callingpolicy','calling')]
	[string]
	$param_callingpolicy
	,
	
	[Alias('package','policypackage')]
	[string]
	$param_policypackage
	,
	
	[Alias('singleuser')]
	[string]$param_singleuser
	,
	
	[Alias('nodefaults')]
	[switch]$param_nodefaults
	,
	
	[Alias('showlog','showlogfile','openlog')]
	[switch]$param_showlog

)

# Top level global variables (Variables are script agnostic but with script specific values)
$VERSION 	= "2.3.1"
$SCRIPTNAME = "Bulk Teams Policy Update Script"
$LOGPATH	= "C:\logs\" #needs trailing slash \
$LOGNAME	= "$(Get-Date -format yyyy-MM-dd_HH-mm-ss)_$($SCRIPTNAME -replace ' ', '').log"
$strGlobalLogDestination = "$LOGPATH$LOGNAME"

#Start Script
Function Main {

	#set the target file for logging
	If ($param_logfilepath) { 
		Set-LogFile $param_logfilepath
		$strGlobalLogDestination = $param_logfilepath
	} 
	If (-Not $param_nodefaults) {
		If (Test-Path($strGlobalLogDestination)) {
			Set-LogFile $strGlobalLogDestination
		} else {
			Log-Entry "No acceptable logfile path provided in paramaters or coded in script defaults. Using default log file location" -foreground 'DarkGray'
			$strGlobalLogDestination = "$env:Temp\$($My.Name).Log"
		}
	}
	Log-Entry "Logging to $strGlobalLogDestination" -foreground 'DarkGray'
	#open the logfile
	If ($param_showlog) { Start-Process powershell.exe "Get-Content $strGlobalLogDestination -wait" }

	
$CHANGELOG_TEXT = "
 ==========================================================================
  -----------------------------Changelog----------------------------------
   0.01 - Initial construction                                 2020-03-11
   2.00 - Added standard scripting framework                   2020-03-13
   2.01 - Improved Email-Report function to include parameters 2020-03-13
   2.10 - Added policy package assignment                      2020-03-13
        - fixed default policy not applying if param empty
		- added -nodefaults (only apply things if explicitly specified)
   2.11 - fixed gui not indicating success on package apply    2020-03-15
        - fixed -nodefault causes errors and crashing          2020-03-15
   2.20 - added logging of each bulk operation                 2020-03-28
        - re-wrote Inform-Operator to use Log-Entry
        - replaced most remaining Write-Host with Log-Entry
        - improved wording of many log entries
   2.2.1 - switched to semantic versioning                     2020-04-01
         - minor syntax best practice corrections
   2.2.2 - removed semi-sensitive domains and paths etc.       2020-04-10
         - improved help and comment text
   2.3.0 - improved logfile default location handling          2020-04-11
   2.3.1 - fixed -showlog not working (introduced in 2.3.0)    2020-04-11
  ------------------------------Credits-----------------------------------
  Various internet sources may be used in the writing of this script.
  Sources and any code copied verbatim, will be noted in the function header
 =========================================================================="
 
# Define the requirements to be output later. This section used to be above.
$REQUIREMENTS_TEXT = "
===========================================================================
                              Requirements: 
---------------------------------------------------------------------------
 - Privileges:
	- [Credentials] O365 Global Admin
 - Powershell Modules
	- MSOnline
	- MicrosoftTeams (Install-Module MicrosoftTeams)
	- Active Directory module (if using AD group or OU as input)
 - Infastructure
    - Access to an SMTP server for sending email reports
 - For more details on usage, use the -help argument.
---------------------------------------------------------------------------"

$FEATURES_TEXT = "
===========================================================================
       LIST OF PLANNED AND IMPLEMENTED FEATURES: [x] = implemented
---------------------------------------------------------------------------
Script Specific:
- [x] Single user mode (-singleusermode) to specify single user      [2.11]
- [x] Take in list of User UPNs from CSV                             [0.10]
- [ ] Check if row is valid user and skip if not
- [ ] Supply credentials via CLI instead of waiting for GUI prompt
- [x] Log success or fail (and errmsg if fail) for each row item     [2.20]
- [ ] Take AD group as user source
- [ ] Take OU as user source

Script Agnostic: (modularise later)
- [x] Switches -requirements -features -help -changelog -bugs        [0.01]
- [x] Comment Based Help [0.01]
- [ ] Switch to attach results as CSV
- [ ] Automatically attach results as CSV if too large
- [ ] List log file path at start and end of log
- [ ] take logfile path at commandline
- [ ] take email report recipient at commandlline
- [ ] take email sender and recipient domain suffix at CLI (-emaildomain)
- [ ] separate params for sender and recipient domain suffixes
- [ ] take -nogui or -silent switches to remove all non-pipeable console output
- [x] Semantic versioning                                            [2.2.1]
 ==========================================================================="
 
$HELP_TEXT = "
===========================================================================
                              Help:
---------------------------------------------------------------------------
 - Switches:
    -help           displays this help text
    -requirements   displays detailed requirements for running the script
    -features       displays list of implemented and planned features
    -changelog      displays history of changes and versions
    -issues         displays list of known issues (also -bugs, -knownissues)
	-test           runs the script in test mode [script specific]
	-showlog        opens the logfile in a new powershell window with gc -wait
	-verbose        shows verbose lines in console (normally only show in logfile)
	-debug			logs debug lines in the logfile (normally not logged)
 - Switches: [script specific]
	-singleuser     runs the script on a single user (takes a UPN)
    -csv            specifies csv to use as input (needs full path)
 - Examples
	- Example1:     "".\$SCRIPTNAME.ps1"" -features
	- Example2:     Get-Help .\$SCRIPTNAME.ps1 -examples
 -------------------------------------------------------------------------
 Other info: Help template version 0.3 [updated from 0.2 on 2020-04-01]
==========================================================================="

$ISSUES_TEXT = "
===========================================================================
                              Known Issues:
    (delete and move to changelog when fixed. first line is an example)
---------------------------------------------------------------------------
[Version     [Bug       [Issue
 introduced]  Index]     Description]
e.g. x.x.x         x.x.x-[a-z] description of issue goes here
---------------------------------------------------------------------------

==========================================================================="

# ++++++++++++++++++++++++++++++++
#  Initialise global variables
# ++++++++++++++++++++++++++++++++
# Script Agnostic Defaults
$arrErrors = "1","1"
$bolGlobalTestState = $param_test
$bolGlobalTestModeRollCall = $false #initialise the variable to false, it should be checked later in the script body section
$strGlobalCurrentDirectory = "TOIMPLEMENT"
$ErrorActionPreference = 'Stop'
$strGlobalDefaultDomain = "$env:USERDNSDOMAIN"
$globalReportRecipients = "ITAlerts@$strGlobalDefaultDomain"
$globalSMTPserver = "mail.$strGlobalDefaultDomain"

# Script Specific Defaults
$strGlobalCallingPolicy = "Disable Calling"
$strGlobalMessagingPolicy = "Students-JSR-2020"
$strGlobalPolicyPackage = "Education_PrimaryStudent"

# Script Specific Parameters

# Declare constants (script agnostic)
$CONST_ARRAY_YESRESPONSES = "y","yes"
$CONST_HORIZONTAL = "horizontal"
$CONST_VERTICAL = "vertical"
$CONST_END = "end"
$CONST_START = "start"
$CONST_FAST = "fast"
$CONST_SLOW = "slow"
$CONST_MEDIUM = "medium"
$CONST_TESTON = $true
$CONST_TESTOFF = $false

<# -------------------------------------
.SYNOPSIS
#  Function to get current line number

.DESCRIPTION
	When called, will give the line number on which it was called.
	Used for debugging
	Credits: Kirk Munro

.VERSION
	1.0

.LINK
https://poshoholic.com/2009/01/19/powershell-quick-tip-how-to-retrieve-the-current-line-number-and-file-name-in-your-powershell-script/
# ------------------------------------- #>
function Get-CurrentLineNumber { 
    $MyInvocation.ScriptLineNumber 
}


<# -------------------------------------
.SYNOPSIS
Display text in the console to inform the operator of something.

.DESCRIPTION
Takes various preset switches and strings and displays predefined text.
Has various options such as success or failure, dots ... to indicate progress etc.
Partially script agnostic (won't do anything unless called, but contains
script specific presets which need to be cleaned up and adjusted for each script)

.VERSION
2.0 - updated for compatibility with Log-Entry

.PARAMETER start
Used in conjunction with -function.
Displays "- Starting <functionname> ..." on the console. The dots are 

.EXAMPLE
PS> Inform-Operator -start -function "Waiting for export"

# ------------------------------------- #>
Function Inform-Operator {

	#-----------------Start standard function header------------------#
	param (	
		[switch]$testfunction
		,
		[ValidateSet("testmode","livemode","waitmode","disablemode","strictmode","noexportmode","initiate","requirements","features","changelog","help","issues")]
		[string]$preset
		,
		[switch]$start
		,
		[switch]$end
		,
		[ValidateSet("success","fail","start","true","false","pending")]
		[string]$state
		,
		[string]$function
		,
		[boolean]$long
		,
		[string]$notes = ""
	)
	
	$return = @{}
	$return["Function"] = $MyInvocation.MyCommand.Name
	$return["Version"] = 1.5
	$return["Result"] = $false #set it to false at first, so if nothing happens to set it as true we consider the function a failure
	$return["LastMessage"] = "This function will almost always return true as there's no way to verify Write-Host calls"
	
	Log-Debug "Entering $($return.function) function, version $($return.Version)"
	#------------------End standard function header--------------------#
	
	#set function variables
	$intDefaultNumDots = 3
	$intDefaultLineLength = 50
	
	If ($bolGlobalTestModeRollCall) {
		$return.Result = $true
		$return.LastMessage = "Breaking from function due to rollcall"
		Return $return
	} elseif ($WhatIfPreference) {
		$return.Result = $true
		Log-Verbose "Whatif switch detected, leaving $($return.function) function, version $($return.Version)"
	} else {
	
		#Script Agnostic
		$strPresetInitiate = "Initiating script. Checking switches and arguments..."
		$strPresetTestMode = "Test mode has been indicated."
		$strPresetLiveMode = "Live mode has been indicated."
		#Script Specific
		$strPresetStrictMode = "Strict checking has been indicated."
		
		#process the preset argument
		Switch ($preset.ToLower()) {
			"testmode" {
				Log-Entry $strPresetTestMode -foreground "DarkGreen" -strip:-1
			}
			"livemode" {
				Log-Entry $strPresetLiveMode -foreground "Black" -background "Yellow" -strip:-1
			}
			"initiate" {
				Log-Verbose $($return.function): detected -$preset parameter from commandline
				Log-Entry $strPresetInitiate -foreground "DarkGray" -strip:-1
			}
			"requirements" {
				Log-Verbose $($return.function): detected -$preset parameter from commandline
				Log-Entry $REQUIREMENTS_TEXT -foreground "Green" -strip:-1
			}
			"features" {
				Log-Verbose $($return.function): detected -$preset parameter from commandline
				Log-Entry $FEATURES_TEXT -foreground "DarkYellow" -strip:-1
			}
			"changelog" {
				Log-Verbose $($return.function): detected -$preset parameter from commandline
				Log-Entry $CHANGELOG_TEXT -foreground "DarkCyan" -strip:-1
			}
			"help" {
				Log-Verbose $($return.function): detected -$preset parameter from commandline
				Log-Entry $HELP_TEXT -foreground "Magenta" -strip:-1
			}
			"issues" {
				Log-Verbose $($return.function): detected -$preset parameter from commandline
				Log-Entry $ISSUES_TEXT -foreground "Magenta" -strip:-1 
			}
		}
		#script specific switches
		Switch ($preset.ToLower()) {
			"strict" {
				Log-Verbose $($return.function): detected -$preset parameter from commandline
				Log-Entry $strPresetStrictMode -foreground "Green"
			}
		}
		
		#for ease of use of the function, I've made start a switch parameter
		#for cleanness of code, I process it in the "Switch" statement below which requires a conversion
		#since -start should never be called at the same time as "success" or "fail" it's safe to do this
		If ($start) { $state = "start" }
		
		#process the state argument
		Switch ($state.ToLower()) {
			"start" {
				Log-Entry "- Starting " -nonewline
				Log-Entry $function -foreground "cyan" -nonewline
				TimedDots -numdots $intDefaultNumDots -direction $CONST_HORIZONTAL -speed $CONST_FAST
				[string]$strEmptyString = ""
				$padding = $intDefaultLineLength-$intDefaultNumDots-$function.Length
				if ($padding -lt 0) { $padding = 30 }
				Log-Entry $strEmptyString.PadRight($padding," ") -nonewline
			}
			"end" {
				If ($long) { Log-Entry "Finished processing function $paramFunctionName."; Log-Entry "" }
			}
			{@("success","true") -contains $_} {
				Log-Entry "success`t" $notes -foreground "green"
			}
			{@("fail","false") -contains $_} {
				Log-Entry "fail`t" $notes -foreground "red"
			}
			"pending" {
				Log-Entry "pending`t" $notes -foreground "darkyellow"
			}
		}
		$return.Result = $true
	}
	
	Write-Debug "Leaving $($return.function) function, version $($return.Version)"
	Write-Debug $($return)
}

# ----------------------------------------
#
#  Function for writing a series of dots
#	
#	Version: 2.0
#
# ----------------------------------------
Function TimedDots {

	#-----------------Start standard function header------------------#
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param (
		[ValidateRange(1,10)]
		[int]$numdots
		,
		[ValidateSet("vertical","horizontal")]
		[string]$direction
		,
		[ValidateSet("slow","medium","fast")]
		[string]$speed
		,
		[switch]$newline
		,
		[ValidateRange(50,200)]
		[int]$speedexpert
	)
	
	#standard function stuff-------------
	$return = @{}
	$return["function"] = $MyInvocation.MyCommand.Name
	$return["version"] = 2.0
	$return["LastMessage"] = "Initalising function"
	
	Write-Debug "Entering $($return.function) function, version $($return.Version)"
	#------------------End standard function header--------------------#
	
	If (-Not $PSCmdlet.ShouldProcess("")) {
		$return.Result = $true
		Write-Debug "Leaving $($return.function) function, version $($return.Version) due to -WhatIf"
	}
	
	#declare variables
	$intSpeed = 0
	$strSecMilliSec = "m"
	
	#set the number of milliseconds depending if user requests fast or slow
	Switch ($speed.ToLower()) {
		"slow" { $intSpeed = 700 }
		"medium" { $intSpeed = 400 }
		"fast" { $intSpeed = 100 }
	}
	
	#write dots with or without newline depending on user request
	Switch ($direction.ToLower()) {
		"horizontal" {
			foreach ($i in 1..$numdots) {
				Write-Host "." -nonewline
				Start-Sleep -m $intSpeed
			}
			If ($newline) { Write-Host "" } #else { Write-Host "`t`t`t" -nonewline }
		}
		"vertical" {
			foreach ($i in 1..$numdots) {
				Write-Host "."
				Start-Sleep -m $intSpeed
			}
		}
	}
	
	$return.Result = $true
	Write-Debug "Leaving $($return.function) function, version $($return.Version)"
	Write-Debug $($return)
}

<# -------------------------------------
.NAME
	Function-Template

.SYNOPSIS
	Template to save time creating a new function by allowing quick copy and paste
	Also assists in maintaining uniformity over all functions and reducing error.
	
.DESCRPTION
	Version:	1.1
	Author:		James Tuson
	References:	n/a

.CHANGELOG
	0.8 - updated Return $true to $return.result = $true (same for $false)		[2018]
		- put start and end comments for prep and action blocks
	0.9 - updated to include parent function and line number in verbose header	[2019-05-xx]
	1.0 - added verification section											[2019-06-01]
	1.1 - added Comment Based Help to the function header						[2019-10-14]
		- Moved the various params onto their own lines and added Alias example

.TODO
	[ ] item to do
#--------------------------------------- #>
Function Function-Template {
	
	#-----------------Start standard function header------------------#
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (	
	[Parameter(Mandatory=$true)][switch]$myswitch,
	[Alias('alias1','alias2')][switch]$alias,
	[string]$somestring
	)
	
	$return = @{}
	$return["function"] = $MyInvocation.MyCommand.Name
	$return["version"] = 0.1
	$return["lastmessage"] = ""
	$return["result"] = $false #set it to false at first, so if nothing happens to set it as true we consider the function a failure
	
	Log-Verbose "Entering $($return.function) function, version $($return.Version)"
	Log-Verbose "I was called by: $($(Get-PSCallStack)[1].Command) from line: $($(Get-PSCallStack)[1].ScriptLineNumber)"
	#------------------End standard function header--------------------#
	
	If ($bolGlobalTestModeRollCall) {
		$return.Result = $true
		Return $return
	} else {
		#------------start of pre-action prep----------------
		
			#replace this with variables, preparation etc.
		
		#-------------end of pre-action prep-----------------
	
		If ($PSCmdlet.ShouldProcess($thingbeingmodified,'Action being done')) {
			Try {
				#------------start of actions------------------
					#action here
				#-------------end of actions-------------------
				#$return.result = $true #uncomment this line once the action is added
			} catch {
				$return.LastMessage = "$_." #this is whatever error just happened
				$return.result = $false
			}
		}
		
		#-------------Verify----------------# [1.0]
		Log-Verbose "Verifying if action succeeded"
		Try {
			$newlyRetrievedValue = "insert function to get value here"
		} catch {
			Log-Verbose "Was not able to get updated value for reason: $($_)"
			If ($bolGlobalStrictChecking) {
				$return.Result = $false
				$return.LastMessage = "Not verified. Error: $_"
				Return $return #bail out since no verify counts as total failure in strict mode
			} else {
				$return.LastMessage = "No errors during action, but unable to verify due to $_"
			}
		}
		#-----compare against expected value
		If ($newlyRetrievedValue -eq $expectedValue) {
			$return.LastMessage = "verified that action succeeded"
			$return.result = $true
		} else {
			$return.result = $false
			$return.LastMessage = "no error during action, but action hasn't applied"
		}
	}
	
	Log-Verbose "Leaving $($return.function) function, version $($return.Version). Result: $($return.result)"
	Return $return
}

# -------------------------------------
#
# Report formatting function
#
#	VERSION
#		1.2
#	USAGE
#		Takes an array of rows and formats it in a way that can be emailed nicely
#	LIMITATIONS
#
#	CHANGELOG
#		0.1 - Adapted from auditlogreport.ps1. Unknown source.
#		0.5 - Removed Begin,Process,End sections, adapted to fit the array it is passed.
#			- Still not working though. When bugs are fixed, will updated to 1.0.
#		1.0 - rewrote to handle any sort of array
#			- fixed ($bodyarray[$i] needed to be $bodyarray[$i-1]
#			- [known bug 1.0] Table header not spanning full table width
#		1.1 - added handling for bodyarray being actual array or just single string [0.68]
#		1.2 - fixed bug from 1.1 which wasn't working. now handles bodyarray being a string
#	TODO
#		[Done-1.0/0.64]Re-write so not tied to specific column names and can be used in any script
#		[Done-0.66]Add global script flags to table header row e.g. waitforexport
#		Fix so function can handle either array or string as input
#
# -------------------------------------
function Format-Report {
    
	[CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $bodyarray

        )
	
	$css = @'
	<style type="text/css">
	body { font-family: Tahoma, Geneva, Verdana, sans-serif;}
	table {border-collapse: separate; background-color: #F2F2F2; border: 3px solid #103E69; caption-side: bottom;}
	td { border:1px solid #103E69; margin: 3px; padding: 3px; vertical-align: top; background: #F2F2F2; color: #000;font-size: 12px;}
	thead th {background: #903; color:#fefdcf; text-align: left; font-weight: bold; padding: 3px;border: 1px solid #990033;}
	th {border:1px solid #CC9933; padding: 3px;}
	tbody th:hover {background-color: #fefdcf;}
	th a:link, th a:visited {color:#903; font-weight: normal; text-decoration: none; border-bottom:1px dotted #c93;}
	caption {background: #903; color:#fcee9e; padding: 4px 0; text-align: center; width: 40%; font-weight: bold;}
	tbody td a:link {color: #903;}
	tbody td a:visited {color:#633;}
	tbody td a:hover {color:#000; text-decoration: none;
	}
	</style>
'@

	#alternative html style, use this once everything is working OK
	<#
	$altcss = 
	"<style>BODY{font-family: Arial; font-size: 10pt;}
	TABLE{border: 1px solid black; border-collapse: collapse;}
	TH{border: 1px solid black; background: #dddddd; padding: 5px; }
	TD{border: 1px solid black; padding: 5px; }
	</style>"
	#>

	$sb = New-Object System.Text.StringBuilder
	[void]$sb.AppendLine($css)
	[void]$sb.AppendLine("<table cellspacing='0'>")
	
	#--------------------
	# table header row 1
	#--------------------
	$span = $(If ($string) { 10 } else { If($bodyarray) { $bodyarray[0].Count } { else { Log-Entry "problem!" } } } )
	#to revert, replace "span" with "bodyarray[0].Count" in the line below
	[void]$sb.AppendLine("
	<tr>
		<td colspan=$($span)>
			<strong>
				<p>$SCRIPTNAME Report for $((get-date).ToShortDateString())
				<br>Script version $VERSION</p>
			</strong>
		</td>
	</tr>
	")
	
	#--------------------
	# table header row 2
	#--------------------
	#delete this section if your script doesn't have global flags
	#to revert, replace "span" with "bodyarray[0].Count"
	[void]$sb.AppendLine("
	<tr>
		<td colspan=$($bodyarray[0].Count)>
			<table>
	")
	$params = (Get-Command -Name $PSCommandPath).Parameters
	Log-Debug $params
	
	foreach ($param in $params.Keys) {
		$paramvalue = (Get-Variable -Name $param -EA SilentlyContinue).Value
		[void]$sb.AppendLine("
				  <tr>
					<td>$param</td>
					<td>$paramvalue</td>
				  </tr>
		")
	}
	[void]$sb.AppendLine("
			</table>
		</td>
	</tr>
	")
	#--------------------
	# table body
	#--------------------
	#check if the line is an array or just a string [function v1.1]
	If (($bodyarray.GetType()).Name -like "String") {
		#if its just a string, just add it into a single row in the table
		[void]$sb.AppendLine("<tr>") #start of row
		[void]$sb.AppendLine("<td>	$($bodyarray)	</td>")
		[void]$sb.AppendLine("</tr>") #end of row
	} else {
		[void]$sb.AppendLine("<tr>")
		ForEach ($key in $bodyarray[0].Keys) {
			[void]$sb.AppendLine("<td><strong>$key</strong></td>")
		}
		[void]$sb.AppendLine("</tr>")
		For ($i=1; $i -le $bodyarray.Count; $i++) {
			[void]$sb.AppendLine("<tr>")
			ForEach ($line in $bodyarray[$i-1].GetEnumerator()) {
				[void]$sb.AppendLine("<td>	$($line.Value)	</td>")
			}
			[void]$sb.AppendLine("</tr>")
		}
	}
	
	#close off the table code
	[void]$sb.AppendLine("</table>")
	
	#return the result
	Write-Output $sb.ToString()
}

# -------------------------------------
#
# Email Report Function
#
#	CHANGELOG
#	0.1 - Made function generic, removed script specific hardcoding
#	0.2 - Added body, subject, recipients, sender as parameters
#	0.3 - Added defaults for these values and associated logic
#	1.0 - Rewrote to use different smtp sending method
#		- Added html formatting option
#		- Bug fixes and corrections, changed argument names
#		- Seperated datesuffix out to account for subject being specified
#	1.1 - Removed html formatting (now handled by Format-Report function)
#	1.2 - Added check if $body is a string to handle empty Leaving Staff
#	1.3 - Renamed function from EmailReport to Email-Report
#	1.4 - updated $from parameter data type to [mailaddress]
#		- reformatted param block, params on their own lines
#		- removed references to other script versions in this header
#	1.5 - Fixed mistake in body html check. Log messages were reversed
#		  for simple string and array.
#	1.6 - Changed email body log level to debug
#		- Removed function name from email attributes
#
#	TODO
#	 Code for multiple recipients
#	 Code to attach log file and csv to email
#	 Fix authentiction issue or allow specification of credentials
#
# -------------------------------------
Function Email-Report {

	#-----------------Start standard function header------------------#
	[cmdletbinding(SupportsShouldProcess=$True)]
	param (	
		$body,
		[string]$subject,
		[string[]]$to,
		[mailaddress]$from,
		[switch]$html
	)
	
	$return = @{}
	$return["function"] = $MyInvocation.MyCommand.Name
	$return["version"] = 1.6
	$return["lastmessage"] = ""
	$return["result"] = $false #set it to false at first, so if nothing happens to set it as true we consider the function a failure
	Log-Verbose "Entering $($return.function) function, version $($return.Version)"
	Log-Verbose "I was called by: $($(Get-PSCallStack)[1].Command) from line: $($(Get-PSCallStack)[1].ScriptLineNumber)"
	#------------------End standard function header--------------------#
	
	#set up function default variables
	#helps in troubleshooting e.g. catching instances where a var isn't being set
	#also saves time troubleshooting errors, ensures at least a basic email gets through
	$datesuffix = "[" + $((get-date).ToShortDateString()) + "]"
	$email = @{}
	$email.sender = "Email-ReportfunctionV" + $return.version + "@$strGlobalDefaultDomain"
	$email.recipients = $globalReportRecipients
	$email.subject = "Generic Email Report Subject " + $datesuffix
	$email.body = "Generic Remail Report Body in email report function version " + $return.version
	$email.smtpserver = $globalSMTPserver
	
	#if parameters were passed, replace the defaults
	If($from) 		{ $email.sender = $from }
	If($to)			{ $email.recipients = $to }
	If($subject) 	{ $email.subject = $subject + $datesuffix }
	If($html)		{ $email.body = $body | ConvertTo-HTML -Head $style }
	Else{ If($body)	{ $email.body = $body } }
	Log-Debug "$($return.function): body: $($body)"
		
	#set up the mssage object
	$messageObject = New-Object System.Net.Mail.MailMessage $email.sender, $email.recipients
	$messageObject.Subject		= $email.subject
	$messageObject.IsBodyHTML	= $true
		
	#check and tell Format-Report if it's just a simple string 
	Log-Verbose '$body.GetType() is:' "$($body.GetType())"
	If (($body.GetType()).Name -ne "String")	{
		Log-Verbose 'We have detected the email message body is an array or other object'
		Log-verbose 'Calling Format-Report and passing it the $body variable'
		$messageObject.Body = Format-Report -bodyarray $body
	} else {
		Log-Verbose 'We think the email message body is a simple string'
		Log-verbose 'Calling Format-Report and passing it the $body variable'
		$messageObject.Body = Format-Report -bodyarray $body
	}
	
	If ($bolGlobalTestModeRollCall) {
		$return["result"] = $true
		Return $return
	} else {
		#send the lastmessage
		Log-Debug "$($return.function): subject: $($email.subject)"
		Log-Debug "$($return.function): body: $($email.body)"

		If ($PSCmdlet.ShouldProcess($email.recipients,"Sending email of subject ""$($email.subject)"" with body ""$($email.body)""")) {
			Try {
				Log-Verbose "To: $($email.Recipients)"
				Log-Verbose "From: $($email.sender)"
				Log-Verbose "Subject: $($email.subject)"
				Log-Debug "Body: $($email.body)"
				Log-Verbose "SmtpServer: $($email.smtpserver)"
				<#Send-MailMessage -To $email.recipients -From $email.sender `
					-Subject $email.subject `
					-Body $email.body `
					-SmtpServer $email.SmtpServer `
					-BodyAsHTML#>
				$smtpObject = New-Object Net.Mail.SmtpClient($email.smtpServer)
				$smtpObject.Send($messageObject)
				$return.Result = $true
			} catch {
				$return.Result = $false
			}
		}
	}
	
	Log-Verbose "Leaving $($return.function) function, version $($return.Version)"
	Return $return
}

#==================================================
#
#  Beginning of script processing
# -----------------------------------------------
#
#	LIST OF ACTIONS: 
#	1. Process the switches and arguments
#	2. Give operator info on what's happening
#	3. Initiate either bulk or single user mode [planned]
#	4. Give operator post-process warning/info
#
#	STRATEGY:
#	1. see $FEATURES_TEXT at top of script
#	2. have each function testable
#	3. modularise actions and link to cli switches
# -----------------------------------------------
#
# =================================================

Write-Host ""
#$null = Inform-Operator -preset "initiate"
Write-Host ""
$date = Get-Date


# Announce version to operator [Script Agnostic]
Write-Host "`n" ("Welcome").Padleft(40," ") "`n" -foreground "Yellow"
Log-Entry "You are running version " -nonewline
Log-Entry $VERSION -foreground "Magenta" -nonewline
Log-Entry " of the " -nonewline
Log-Entry $SCRIPTNAME -foreground "Green" -nonewline
Log-Entry " script. The time is $date"

Log-Verbose "Now checking switches and arguments"
# Check switches and arguments here
	
	#info dumps
	$exitflag = $false
	If ($requirements) { $null = Inform-Operator -preset "requirements"; $exitflag = $true }
	If ($param_changelog) { $null = Inform-Operator -preset "changelog"; $exitflag = $true }
	If ($param_features) { $null = Inform-Operator -preset "features"; $exitflag = $true }
	If ($param_help) { $null = Inform-Operator -preset "help"; $exitflag = $true }
	If ($param_issues) { $null = Inform-Operator -preset "issues"; $exitflag = $true }
	If ($param_requirements) { $null = Inform-Operator -preset "requirements"; $exitflag = $true }
	If ($exitflag) { Log-Verbose "Exiting due to exitflag"; End }
	
	#script specific
	

Log-Entry ""
Log-Entry ' +-----------------------------------------------------------------------------------------+' -foreground Green
Log-Entry ' | Note: this script assumes you have already established connections to O365 and MS Teams |' -foreground Green
Log-Entry ' | via the Connect-MSOLService and Connect-MicrosoftTeams cmdlets. May automate this later.|' -foreground Green
Log-entry ' | This script uses a modified version of Log-Entry ( https://github.com/iRon7/Log-Entry ) |' -foreground Green
Log-Entry ' +-----------------------------------------------------------------------------------------+' -foreground Green
Log-Entry ' A report will be emailed to $globalReportRecipients (' -nonewline
Log-Entry "$globalReportRecipients" -nonewline -foreground magenta; Log-Entry '). Email recipient can be specified using the  -emailrecipient parameter'
Log-Entry ""

#set the array counter. Subtract one because powershell array indexes start at 0
$counter = $param_resumeat - 1

If (!$param_singleuser) {
	#if the csv input path hasn't been specified on the commandline, prompt the user for it
	if (!$param_csvpath) {
		Log-Verbose Asking user for CSV path
		$csvpath = read-host -prompt "enter path to csv"
	} else {
		$csvpath = $param_csvpath
	}

	Try {
		$users = import-csv $csvpath
	} catch {
		Log-Entry 'Failed to import CSV. Reason as follows:'
		Write-Host $_
		Log-Verbose $_
		End #quit since no users to process
	}
} else {
	$users = @()
	$userrow = @{}
	$userrow["mail"] = $param_singleuser
	$users += $userrow
}

#set up row template. 
#this template is copied and each instance added to $resultsArray
$rowTemplate = @{
	    counter = "0"
		    UPN = "-"
 samaccountname = "-"
	 callresult = "-"
	  callerror = "-"
  messageresult = "-"
   messageerror = "-"
  packageresult = "-"
   packageerror = "-"
}

$resultsArray = @()

$callingToApply		= $(If($param_callingpolicy) {$param_callingpolicy} else { If($param_nodefaults) { $null } else { $strGlobalCallingPolicy }})
$messagingToApply	= $(If($param_messagingpolicy) {$param_messagingpolicy} else { If($param_nodefaults) { $null } else { $strGlobalMessagingPolicy }})
$packageToApply		= $(If($param_policypackage) {$param_policypackage} else { If($param_nodefaults) { $null } else { $strGlobalPolicyPackage }})

If ($callingToApply -or $messagingToApply -or $packageToApply) {
	do { 

		$row = $rowTemplate.PSObject.Copy()
		$row.counter	= $counter+1
		$row.UPN		= $users[$counter].Mail
		$row.username	= $users[$counter].SamAccountName
		$guicounter = $counter+1
		
		$UPN = $users[$counter].Mail
		$name = $users[$counter].GivenName + " " + $users[$counter].Surname
		
		Inform-Operator -start -function "$guicounter of $($users.count) $($row.username) $name $UPN"
		
		#Write-Host "Processing $guicounter of $($users.count) " -nonewline
		#Write-Host " $($users[$counter].Mail)" -nonewline
		
		#check if policypackage, if so don't do calling or messaging
		If (!$param_policypackage) {
			#apply the calling policy
			Try {
				Grant-CsTeamsCallingPolicy -PolicyName $callingToApply -Identity $UPN -EA SilentlyContinue
				$counter++
				$row.callresult = "OK"
			} catch {
				$counter++
				$row.callresult = "fail"
				$row.callreason = $_
				#Log-Entry " $($row.Callresult)" -foreground Red
			}
			
			#apply the messaging policy
			Try {
				Grant-CsTeamsMessagingPolicy -PolicyName $messagingToApply  -Identity $UPN -EA SilentlyContinue
				#Log-Entry " Done" -foreground Green
				$counter++
				$row.messageresult = "OK"
			} catch {
				#Log-Entry " Fail" -foreground Red
				$counter++
				$row.messageresult = "fail"
				$row.messagereason = $_
			}
			#check if either of the calls have an error present (i.e. they failed)
			If (!($row.callreason -or $row.messagereason)) {
				Inform-Operator -state "success"
			} else {
				Inform-Operator -state "fail" -notes "$($row.messagereason) $($row.callreason)"
			}
		} else {
			#apply the policy package
			Try {
				Grant-CsUserPolicyPackage -PackageName $packageToApply -identity $UPN -EA SilentlyContinue
				$counter++
				$row.packageresult = "OK"
				Inform-Operator -state "success"
			} catch {
				#Log-Entry " Fail" -foreground Red
				$counter++
				$row.packageresult = "fail"
				$row.packagereason = $_
				Inform-Operator -state "fail" -notes $($row.packagereason)
			}
		}
		
		$resultsArray+=$row
		$row = $null
		
	} while ($counter -le ($users.Count-1))
	Log-Verbose 'finished processing all items.'
} else {
	$message = "$($users.Count) of $($users.Count) users skippped as there was nothing to apply"
	Log-Entry $message
	$resultsArray = $message
}

#email report
Function InternalOnly-EmailReport {
	Write-Host ""
	Inform-Operator -start -function "Emailing report"
	Log-Verbose "Turn on debug logging (-debug) to view the array being sent to the email report function" -nonewline
	Log-Debug $resultsArray -expand
	$emailResult = Email-Report -body $resultsArray -html -from "BulkOperationScript@$strGlobalDefaultDomain" -subject "Bulk Operations at $(get-date -format HH:mm)" -to $globalReportRecipients
	Inform-Operator -state $emailResult.Result -notes $emailResult.LastMessage
}
InternalOnly-EmailReport

# Stuff for end
Log-Entry ""
Log-Entry "Please perform manual checks to verify all changes were completed accurately." -foreground "yellow"
Log-Entry "Responsibility for verifying accuracy is on the operator of the script." -foreground "yellow"
Log-Entry ""
Log-Entry "If appropriate, proceed with remaining manual steps now."
Log-Entry ""

} #end of Main function


# ------------------------------------- Global --------------------------------
Function Global:ConvertTo-Text([Alias("Value")]$O, [Int]$Depth = 9, [Switch]$Type, [Switch]$Expand, [Int]$Strip = -1, [String]$Prefix, [Int]$i) {
	Function Iterate($Value, [String]$Prefix, [Int]$i = $i + 1) {ConvertTo-Text $Value -Depth:$Depth -Strip:$Strip -Type:$Type -Expand:$Expand -Prefix:$Prefix -i:$i}
	$NewLine, $Space = If ($Expand) {"`r`n", ("`t" * $i)} Else{$Null}
	If ($Null -eq $O) {$V = '$Null'} Else {
		$V = If ($O -is "Boolean")  {"`$$O"}
		ElseIf ($O -is "String") {If ($Strip -ge 0) {'"' + (($O -Replace "[\s]+", " ") -Replace "(?<=[\s\S]{$Strip})[\s\S]+", "...") + '"'} Else {"""$O"""}}
		ElseIf ($O -is "DateTime") {$O.ToString("yyyy-MM-dd HH:mm:ss")} 
		ElseIf ($O -is "ValueType" -or (($O.Value.GetTypeCode -and $O.ToString.OverloadDefinitions) -and -not $O.GetEnumerator.OverloadDefinitions)) {$O.ToString()}
		ElseIf ($O -is "Xml") {(@(Select-XML -XML $O *) -Join "$NewLine$Space") + $NewLine}
		ElseIf ($i -gt $Depth) {$Type = $True; "..."}
		ElseIf ($O -is "Array") {"@(", @(&{For ($_ = 0; $_ -lt $O.Count; $_++) {Iterate $O[$_]}}), ", ", ")"}
		ElseIf ($O.GetEnumerator.OverloadDefinitions) {"@{", @(ForEach($_ in $O.Keys) {Iterate $O.$_ "$_ = "}), "; ", "}"}
		ElseIf ($O.PSObject.Properties -and !$O.value.GetTypeCode) {"{", @(ForEach($_ in $O.PSObject.Properties | Select-Object -Exp Name) {Iterate $O.$_ "$_`: "}), "; ", "}"}
		Else {$Type = $True; "?"}}
	If ($Type) {$Prefix += "[" + $(Try {$O.GetType()} Catch {$Error.Remove($Error[0]); "$Var.PSTypeNames[0]"}).ToString().Split(".")[-1] + "]"}
	"$Space$Prefix" + $(If ($V -is "Array") {
		$V[0] + $(If ($V[1]) {
			If ($NewLine) {$V[2] = $NewLine}
			$NewLine + ($V[1] -Join $V[2]) + $NewLine + $Space
		} Else {""}) + $V[3]
	} Else {$V})
}; Set-Alias CText ConvertTo-Text -Scope:Global -Description "Convert value to readable text"

Function Global:Log-Entry {
	Param(
		$0, $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15,	# PSv2 doesn't support PositionalBinding
		[ConsoleColor]$BackgroundColor, [Alias("Color")][ConsoleColor]$ForegroundColor, [String]$Separator = " ", [Switch]$NoNewline,
		[Int]$Indent = 0, [Int]$Strip = 120, [Switch]$QuoteString, [Int]$Depth = 1, [Switch]$Expand, [Switch]$Type, [Switch]$FlushErrors,
		[Alias('line','showline','showlinenumber')][Switch]$param_ShowLineNumbersInConsole #v02.01.08
	)
	$Noun = ($MyInvocation.InvocationName -Split "-")[-1]
	$LineNumberOfCallingCode = "{0:0000}" -f $MyInvocation.ScriptLineNumber #v02.01.08
	Function IsQ($Item) {If ($Item -is [String]) {$Item -eq "?"} Else {$False}}
	$Arguments = $MyInvocation.BoundParameters
	If (!$My.Log.ContainsKey("Location")) {Set-LogFile "$Env:Temp\$($My.Name).log"}
	If (!$My.Log.ContainsKey("Buffer")) {
		$My.Log.ProcessStart = Get-Date ((Get-Process -id $PID).StartTime); $My.Log.ScriptStart = Get-Date
		$My.Log.Buffer  = (Get-Date -Format "yyyy-MM-dd") + " `tPowerShell version: $($PSVersionTable.PSVersion), process start: " + (ConvertTo-Text $My.Log.ProcessStart) + "`r`n"
		$My.Log.Buffer += (Get-Date -Format "HH:mm:ss.ff") + "`t$($My.Name) version: $($My.Version), command line: $($My.Path) $($My.Arguments)`r`n"}
	If ($FlushErrors) {$My.Log.ErrorCount = $Error.Count} ElseIf (!$My.Log.ContainsKey("ErrorCount")) {$My.Log.ErrorCount = 0}
	While ($My.Log.ErrorCount -lt $Error.Count) {
		$Err = $Error[$Error.Count - ++$My.Log.ErrorCount]
		$My.Log.Buffer += @("`r`n")[!$My.Log.Inline] + "Error at $($Err.InvocationInfo.ScriptLineNumber),$($Err.InvocationInfo.OffsetInLine): $Err`r`n"}
	If ($My.Log.Inline) {$Items = @("")} Else {$Items = @()}
	For ($i = 0; $i -le 15; $i++) {
		If ($Arguments.ContainsKey("$i")) {$Argument = $Arguments.Item($i)} Else {$Argument = $Null}
		If ($i) {
			$Text = ConvertTo-Text $Value -Type:$Type -Depth:$Depth -Strip:$Strip -Expand:$Expand
			If ($Value -is [String] -And !$QuoteString) {$Text = $Text -Replace "^""" -Replace """$"}
		} Else {$Text = $Null}
		If (IsQ($Argument)) {$Value} Else {If (IsQ($Value)) {$Text = $Null}}
		If ($Text) {$Items += $Text}
		If ($Arguments.ContainsKey("$i")) {$Value = $Argument} Else {Break}
	}
	If ($Arguments.ContainsKey("0") -And ($Noun -ne "Debug" -or $Script:Debug)) {
		$Tabs = "`t" * $Indent; $Line = $Tabs + (($Items -Join $Separator) -Replace "`r`n", "`r`n$Tabs")
		If (!$My.Log.Inline) {$My.Log.Buffer += (Get-Date -Format "HH:mm:ss.ff") + "`t$LineNumberOfCallingCode" + "`t$Tabs"} #modified v02.01.08
		If ($param_ShowLineNumbersInConsole) { $Line = "$LineNumberOfCallingCode`t"+$Line } #v02.01.08
		$My.Log.Buffer += $Line -Replace "`r`n", "`r`n           `t$Tabs"
		If ($Noun -ne "Verbose" -or $Script:Verbose) {
			$Write = "Write-Host `$Line" + $((Get-Command Write-Host).Parameters.Keys | Where {$Arguments.ContainsKey($_)} | ForEach {" -$_`:`$$_"})
			Invoke-Command ([ScriptBlock]::Create($Write))
		}
	} Else {$NoNewline = $False}
	$My.Log.Inline = $NoNewline
	If (($My.Log.Location -ne "") -And $My.Log.Buffer -And !$NoNewline) {
		If ((Add-content $My.Log.Location $My.Log.Buffer -ErrorAction SilentlyContinue -PassThru).Length -gt 0) {$My.Log.Buffer = ""}
	}
}; Set-Alias Write-Log Log-Entry -Scope:Global
Set-Alias Log          Log-Entry -Scope:Global -Description "Displays and records cmdlet processing details in a file"
Set-Alias Log-Debug    Log-Entry -Scope:Global -Description "By default, the debug log entry is not displayed and not recorded, but you can display it by changing the common -Debug parameter."
Set-Alias Log-Verbose  Log-Entry -Scope:Global -Description "By default, the verbose log entry is not displayed, but you can display it by changing the common -Verbose parameter."

Function Global:Set-LogFile([Parameter(Mandatory=$True)][IO.FileInfo]$Location, [Int]$Preserve = 100e3, [String]$Divider = "") {
	$MyInvocation.BoundParameters.Keys | ForEach {$My.Log.$_ = $MyInvocation.BoundParameters.$_}
	If ($Location) {
		If ((Test-Path($Location)) -And $Preserve) {
			$My.Log.Length = (Get-Item($Location)).Length 
			If ($My.Log.Length -gt $Preserve) {									# Prevent the log file to grow indefinitely
				$Content = [String]::Join("`r`n", (Get-Content $Location))
				$Start = $Content.LastIndexOf("`r`n$Divider`r`n", $My.Log.Length - $Preserve)
				If ($Start -gt 0) {Set-Content $Location $Content.SubString($Start + $Divider.Length + 4)}
			}
		}
		If ($My.Log.Length -gt 0) {Add-content $Location $Divider}
	}
}; Set-Alias LogFile Set-LogFile -Scope:Global -Description "Redirects the log file to a custom location"

Function End-Script([Switch]$Exit, [Int]$ErrorLevel) {
	Log "End" -NoNewline
	Log ("(script run time: " + ((Get-Date) - $My.Log.ScriptStart) + ", process run time: " + ((Get-Date) - $My.Log.ProcessStart) + ")")
	If ($Exit) {Exit $ErrorLevel} Else {Break Script}
}; Set-Alias End End-Script -Scope:Global -Description "Logs the remaining entries and errors and end the script"

$Error.Clear()
Set-Variable -Option ReadOnly -Force My @{
	File = Get-ChildItem $MyInvocation.MyCommand.Path
	Contents = $MyInvocation.MyCommand.ScriptContents
	Log = @{}
}
If ($My.Contents -Match '^\s*\<#([\s\S]*?)#\>') {$My.Help = $Matches[1].Trim()}
[RegEx]::Matches($My.Help, '(^|[\r\n])\s*\.(.+)\s*[\r\n]|$') | ForEach {
	If ($Caption) {$My.$Caption = $My.Help.SubString($Start, $_.Index - $Start)}
	$Caption = $_.Groups[2].ToString().Trim()
	$Start = $_.Index + $_.Length
}
$My.Title = $My.Synopsis.Trim().Split("`r`n")[0].Trim()
$My.Id = (Get-Culture).TextInfo.ToTitleCase($My.Title) -Replace "\W", ""
$My.Notes -Split("\r\n") | ForEach {$Note = $_ -Split(":", 2); If ($Note.Count -gt 1) {$My[$Note[0].Trim()] = $Note[1].Trim()}}
$My.Path = $My.File.FullName; $My.Folder = $My.File.DirectoryName; $My.Name = $My.File.BaseName
$My.Arguments = (($MyInvocation.Line + " ") -Replace ("^.*\\" + $My.File.Name.Replace(".", "\.") + "['"" ]"), "").Trim()
$Script:Debug = $MyInvocation.BoundParameters.Debug.IsPresent; $Script:Verbose = $MyInvocation.BoundParameters.Verbose.IsPresent
$MyInvocation.MyCommand.Parameters.Keys | Where {Test-Path Variable:"$_"} | ForEach {
	$Value = Get-Variable -ValueOnly $_
	If ($Value -is [IO.FileInfo]) {Set-Variable $_ -Value ([Environment]::ExpandEnvironmentVariables($Value))}
}

#run the script
Main

End	
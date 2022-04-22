<#
PowerShell Pretty Printer V1 by Dan Ward (dtward@gmail.com)
See dtwconsulting.com or DansPowerShellStuff.blogspot.com for more 
information. And stay tuned for Pretty Printer version 2!

Main function is Edit-DTWCleanScript:

.SYNOPSIS
Cleans PowerShell script: indents code blocks, cleans/rearranges all 
whitespace, replaces aliases with commands, etc.
.DESCRIPTION
Cleans PowerShell script: indents code blocks, cleans/rearranges all 
whitespace, replaces aliases with commands, replaces parameter names with
proper casing, fixes case for [types], etc.  More specifically it:
 - properly indents code inside {}, [], () and $() groups
 - replaces aliases with the command names (dir -> Get-ChildItem)
 - fixes command name casing (get-childitem -> Get-ChildItem)
 - fixes parameter name casing (Test-Path -path -> Test-Path -Path)
 - fixes [type] casing
     changes all PowerShell shortcuts to lower ([STRING] -> [string])
     changes other types ([system.exception] -> [System.Exception]
       only works for types loaded into memory
 - cleans/rearranges all whitespace within a line
     many rules - see Test-AddSpaceFollowingToken to tweak

*****IMPORTANT NOTE: before running this on any script make sure you back
it up, commit any changes you have or run this on a copy of your script.
In the event this script screws up or the get encoding is incorrect (say
you have non-ANSI characters at the end of a long file with no BOM), I 
don't want your script to be damaged!

This version has been tested in PowerShell V2 and in V3 CTP 2.  Let me
know if you encounter any issues.

For all the alias and casing replacement rules above, the function works
on all items in memory, so it cleans your scripts using your own custom
function/parameter names and aliases as well.  The module caches all the
commands, aliases and parameter names when the module first loads.  If
you've added new commands to memory since loading the pretty printer 
module, you may want to reload it.

Also, this uses two spaces as an indent step.  To change this, edit the
line: [string]$script:IndentText = "  "

There are many rules when it comes to adding whitespace or not (see 
Test-AddSpaceFollowingToken).  Feel free to tweak this code and let me
know what you think is wrong or, at the very least, what should be 
configurable.  In version 2 I will expose configuration settings for
the items people feel strongly about.

This pretty printer module doesn't do everything - it's version 1.  
Version 2 (using PowerShell 3's new tokenizing/AST functionality) should
allow me to fix most of the deficiencies. But just so you know, here's
what it doesn't do:
 - change location of group openings, say ( or {, from same line to new
   line and vice-versa;
 - expand param names (Test-Path -Inc -> Test-Path -Include);
 - offer user options to control how space is reformatted.

.PARAMETER SourcePath
Path to the source PowerShell file
.PARAMETER DestinationPath
Path to write reformatted PowerShell.  If not specified rewrites file
in place.
.PARAMETER Quiet
If specified, does not output status text.
.EXAMPLE
Edit-DTWCleanScript -Source c:\P\S1.ps1 -Destination c:\P\S1_New.ps1
Gets content from c:\P\S1.ps1, cleans and writes to c:\P\S1_New.ps1
.EXAMPLE
Edit-DTWCleanScript -SourcePath c:\P\S1.ps1
Writes cleaned script results back into c:\P\S1.ps1
.EXAMPLE
dir c:\CodeFiles -Include *.ps1,*.psm1 -Recurse | Edit-DTWCleanScript
For each .ps1 and .psm1 file, cleans and rewrites back into same file
#>

#region Module and process initialize
# Ensure best practices for variable use, function calling, null property access, etc.
# must be done at module script level, not inside Initialize, or will only be function scoped
Set-StrictMode -Version 2

<#
.SYNOPSIS
Initialize the module: load values of load lookup tables.
.DESCRIPTION
Initialize the module: load values of load lookup tables.
This function is called at end of module definition in the 'main'.
#>
function Initialize-Module {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    # initialize lookup tables
    [hashtable]$script:ValidCommandNames = $null
    [hashtable]$script:ValidCommandParameterNames = $null
    [hashtable]$script:ValidAttributeNames = $null
    [hashtable]$script:ValidMemberNames = $null
    [hashtable]$script:ValidVariableNames = $null
    # load initial values for lookup tables
    Set-LookupTableValues
  }
}


<#
.SYNOPSIS
Initialize the module-level variables used in processing a file.
.DESCRIPTION
Initialize the module-level variables used in processing a file.
#>
function Initialize-ProcessVariables {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    # keep track if everything initialized correctly before allowing script to run,
    # most notably to make sure required modules are loaded; set to true at end of
    # initialize and checked at beginning of main function
    [bool]$script:ModuleInitialized = $false

    # make sure required functions are loaded
    # Get-DTWFileEncoding is required for get encoding function
    [string]$RequiredFunctionName = "DTW.PS.FileSystem"
    $RequiredFunction = Get-Command -Name Get-DTWFileEncoding -ErrorAction SilentlyContinue
    if ($null -eq $RequiredFunction) {
      Write-Error -Message "$($MyInvocation.MyCommand.Name):: Required function $RequiredFunctionName is not loaded; cannot process files."
      return
    }

    #region Initialize process variables
    # indent text; use two spaces by default, change to `t for tabs
    [string]$script:IndentText = "  "

    # initialize file path information
    # source file to process
    [string]$script:PathSource = $null
    # destination file; this value is different from PathSource if Edit-DTWCleanScript -DestinationPath is specified
    [string]$script:PathDestination = $null
    # result content is created in a temp file; if no errors this becomes the result file
    [string]$script:PathDestinationTemp = $null

    # initialize source script storage
    [string[]]$script:SourceScriptStringArray = $null
    [string]$script:SourceScriptString = $null
    [System.Text.Encoding]$script:SourceFileEncoding = $null
    [System.Management.Automation.PSToken[]]$script:SourceTokens = $null

    # initialize destination storage
    [System.IO.StreamWriter]$script:DestinationStreamWriter = $null
    #endregion

    # if got this far, everything is good
    $script:ModuleInitialized = $true
  }
}
#endregion

#region Lookup table functions
#region Set lookup table values in module-level variables

<#
.SYNOPSIS
Populates the values of the lookup tables.
.DESCRIPTION
Populates the values of the lookup tables using the Get-Valid*Names functions.
#>
function Set-LookupTableValues {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    $script:ValidCommandNames = Get-ValidCommandNames
    $script:ValidCommandParameterNames = Get-ValidCommandParameterNames
    $script:ValidAttributeNames = Get-ValidAttributeNames
    $script:ValidMemberNames = Get-ValidMemberNames
    $script:ValidVariableNames = Get-ValidVariableNames
  }
}
#endregion

#region Get initial values for lookup tables

#region Function: Get-ValidCommandNames

<#
.SYNOPSIS
Gets lookup hashtable of existing cmdlets, functions and aliases.
.DESCRIPTION
Gets lookup hashtable of existing cmdlets, functions and aliases.  Specifically
it gets every cmdlet and function name and creates a hashtable entry with the
name as both the key and value; for aliases the key is the alias and the value
is the command name.
When you look up an alias (by accessing the hashtable by alias name for the key,
it returns the command name.  For functions and cmdlet, it returns the same value
BUT the value is the name with the correct case but the lookup is case-insensitive.
Lastly, when the aliases -> are gathered, each alias definition is checked to make
sure it isn't another alias but an actual cmdlet or function name (in the event 
that you have an alias that points to an alias and so on).
#>
function Get-ValidCommandNames {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    [hashtable]$CommandNames = @{}
    # get list of all cmdlets and functions in memory, however we want to sort the 
    # commands by CommandType Descending so Cmdlets come first.  The reason we want
    # Cmdlet to come first is so they get recorded first; in the event that you have
    # a proxy function (function with same name as a cmdlet), you want to use the 
    # value of the Cmdlet name as the lookup value, not the proxy function.
    Get-Command -CommandType Cmdlet,Function | Sort-Object -Property CommandType -Descending | ForEach-Object {
      # only add if doesn't already exist (might not if proxy functions exist)
      if (!($CommandNames.ContainsKey($_.Name))) {
        $CommandNames.($_.Name) = $_.Name
      }
    }

    # for each alias, check its definition and loop until definition command type isn't an
    # alias; then add with original Alias name as key and 'final' definition as value
    Get-Alias d* | ForEach-Object {
      $OriginalAlias = $_.Name
      $Cmd = $_.Definition
      while ((Get-Command -Name $Cmd).CommandType -eq "Alias") { $Cmd = (Get-Command -Name $Cmd).Definition }
      # add alias name as key and definition as value
      $CommandNames.Item($OriginalAlias) = $Cmd
    }

    # last but not least, we want the ForEach Command (i.e. the ForEach-Object Cmdlet used in a 
    # pipeline, not to be confused with the foreach Keyword) to map to the name ForEach-Object.
    # So, let's add it.
    $CommandNames."foreach" = "ForEach-Object"
    $CommandNames
  }
}
#endregion

#region Function: Get-ValidCommandParameterNames

<#
.SYNOPSIS
Gets lookup hashtable of existing parameter names on cmdlets and functions.
.DESCRIPTION
Gets lookup hashtable of existing parameter names on cmdlets and functions. 
Specifically it gets a unique list of the parameter names on every cmdlet and 
function name and creates a hashtable entry with the name as both the key and 
value. When you look up a value, it essentially returns the same value BUT 
the value is the name with the correct case but the lookup is case-insensitive.
#>
function Get-ValidCommandParameterNames {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    [hashtable]$CommandParameterNames = @{}
    #region Parameter casing dilemma
    # Attempting to determine the actual parameter for the command is 'too difficult'
    # with PowerShell version 2's parser token results.  (Stay tuned for V2 of the 
    # pretty printer which will use PowerShell v3 AST/tokenizers...)
    # So, here's the dilemma: with this simple implementation, I'd like to have a 
    # single list of all valid parameter names with the correct casing.  But, as it turns
    # out, I'm seeing different casing between, say Microsoft's parameters and the ones
    # in PowerShell community extensions (NoNewLine vs. NoNewline, Force vs. force, etc.).
    # The Microsoft ones should be the standard, so let's add them first to the hashtable.
    #endregion
    # get list of all unique parameter names on all cmdlets and functions in memory
    # get Microsoft Cmdlet's first, as they should be the reference list for correct letter case
    $Params = Get-Command -CommandType Cmdlet | Where-Object { $_.ModuleName.StartsWith("Microsoft.") } | Where-Object { $null -ne $_.Parameters } | ForEach-Object { $_.Parameters.Keys } | Select-Object -Unique | Sort-Object
    $Name = $null
    $Params | ForEach-Object {
      # param name appears with - in front
      $Name = "-" + $_
      # for each param, add to hash table with name as both key and value
      $CommandParameterNames.Item($Name) = $Name
    }
    # now get all params for cmdlets and functions; the Microsoft ones will already be in
    # the hashtable; add other ones not found yet
    $Params = Get-Command -CommandType Cmdlet,Function | Where-Object { $null -ne $_.Parameters } | ForEach-Object { $_.Parameters.Keys } | Select-Object -Unique | Sort-Object
    $Name = $null
    $Params | ForEach-Object {
      # param name appears with - in front
      $Name = "-" + $_
      # if doesn't exist, add to hash table with name as both key and value
      if (!$CommandParameterNames.Contains($Name)) {
        $CommandParameterNames.Item($Name) = $Name
      }
    }
    $CommandParameterNames
  }
}
#endregion

#region Function: Get-ValidAttributeNames

<#
.SYNOPSIS
Gets lookup hashtable of known valid attribute names.
.DESCRIPTION
Gets lookup hashtable of known valid attribute names.  Attributes
as created by the PSParser Tokenize method) include function parameter 
attributes.
#>
function Get-ValidAttributeNames {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    [hashtable]$AttributeNames = @{
      #region Values for parameter attributes
      <#
      Values for properties on parameter; hard-code for now with correct case 
      until find easy/fast way to programmatically load; using this as the source list
      (classes named *Attribute but no base classes):
      http://msdn.microsoft.com/en-us/library/system.management.automation.aspx
      #>
      Alias = "Alias";
      AllowEmptyCollection = "AllowEmptyCollection";
      AllowEmptyString = "AllowEmptyString";
      AllowNull = "AllowNull";
      CmdletBinding = "CmdletBinding";
      ConfirmImpact = "ConfirmImpact";
      CredentialAttribute = "CredentialAttribute";
      DefaultParameterSetName = "DefaultParameterSetName";
      OutputType = "OutputType";
      Parameter = "Parameter";
      PSDefaultValue = "PSDefaultValue";
      PSTypeName = "PSTypeName";
      SupportsShouldProcess = "SupportsShouldProcess";
      SupportsWildcards = "SupportsWildcards";
      ValidateCount = "ValidateCount";
      ValidateLength = "ValidateLength";
      ValidateNotNull = "ValidateNotNull";
      ValidateNotNullOrEmpty = "ValidateNotNullOrEmpty";
      ValidatePattern = "ValidatePattern";
      ValidateRange = "ValidateRange";
      ValidateScript = "ValidateScript";
      ValidateSet = "ValidateSet";
      #endregion
    }
    $AttributeNames
  }
}
#endregion

#region Function: Get-ValidMemberNames

<#
.SYNOPSIS
Gets lookup hashtable of known valid member names.
.DESCRIPTION
Gets lookup hashtable of known valid member names.  Members (Member tokens
as created by the PSParser Tokenize method) include function parameter 
properties as well as methods on objects.  Unfortunately, there isn't a good
way to get a reasonable subset of the true list of valid values (a list of 
every valid .NET methods?  Yikes!) so these are hard-coded.
This list should be updated as necessary.
#>
function Get-ValidMemberNames {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    [hashtable]$MemberNames = @{
      #region Values for properties on parameter
      <#
      Values for properties on parameter; hard-code for now with correct case 
      until find easy/fast way to programmatically load; using this as the source list:
      http://msdn.microsoft.com/en-us/library/system.management.automation.parameterattribute_members(v=VS.85).aspx
      #>
      HelpMessage = "HelpMessage";
      HelpMessageBaseName = "HelpMessageBaseName";
      HelpMessageResourceId = "HelpMessageResourceId";
      Mandatory = "Mandatory";
      ParameterSetName = "ParameterSetName";
      Position = "Position";
      TypeId = "TypeId";
      ValueFromPipeline = "ValueFromPipeline";
      ValueFromPipelineByPropertyName = "ValueFromPipelineByPropertyName";
      ValueFromRemainingArguments = "ValueFromRemainingArguments";
      #endregion
      #region Oft-used .NET method names and hash table Keys (esp. with splatting)
      Append = "Append";
      Encoding = "Encoding";
      Force = "Force";
      IndexOf = "IndexOf";
      Keys = "Keys";
      NewGuid = "NewGuid";
      Substring = "Substring";
      ToString = "ToString";
      Values = "Values";
      #endregion
    }
    $MemberNames
  }
}
#endregion

#region Function: Get-ValidVariableNames

<#
.SYNOPSIS
Gets lookup hashtable of known valid variables names.
.DESCRIPTION
Gets lookup hashtable of known valid variable names with the correct case.  
It is seeded with some well known values (true, false, etc.) but will grow
as the parser walks through the script, adding user variables as they are 
encountered.
This list could be updated with other known values.
#>
function Get-ValidVariableNames {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    [hashtable]$VariableNames = @{
      #region Values for known variables
      true = "true";
      false = "false";
      HOME = "HOME";
    }
    $VariableNames
  }
}
#endregion

#endregion

#region Loop up value in lookup tables

#region Function: Lookup-ValidCommandName

<#
.SYNOPSIS
Retrieves command name with correct casing, expands aliases
.DESCRIPTION
Retrieves the 'proper' command name for aliases, cmdlets and functions.  When
called with an alias, the corresponding command name is returned.  When called
with a command name, the name of the command as defined in memory (and stored 
in the lookup table) is returned, which should have the correct case.  
If Name is not found in the ValidCommandNames lookup table, it is added and 
returned as-is.  That means the first instance of the command name that is 
encountered becomes the correct version, using its casing as the clean version.
.PARAMETER Name
The name of the cmdlet, function or alias
.EXAMPLE
Lookup-ValidCommandName -Name dir
Returns: Get-ChildItem
.EXAMPLE
Lookup-ValidCommandName -Name GET-childitem
Returns: Get-ChildItem
.EXAMPLE
Lookup-ValidCommandName -Name FunNotFound; Lookup-ValidCommandName -Name funNOTfound
Returns: FunNotFound, FunNotFound
#>
function Lookup-ValidCommandName {
  #region Function parameters
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [string]$Name
  )
  #endregion
  process {
    # look up name in lookup table and return
    # if not found (new function added within script), add to list and return
    if ($ValidCommandNames.ContainsKey($Name)) {
      $ValidCommandNames.Item($Name)
    } else {
      $script:ValidCommandNames.Add($Name,$Name) > $null
      $Name
    }
  }
}
#endregion

#region Function: Lookup-ValidCommandParameterName

<#
.SYNOPSIS
Retrieves command parameter name with correct casing
.DESCRIPTION
Retrieves the proper command parameter name using the command parameter 
names currently found in memory (and stored in the lookup table).
If Name is not found in the ValidCommandParameterNames lookup table, it is
added and returned as-is.  That means the first instance of the command 
parameter name that is encountered becomes the correct version, using its 
casing as the clean version.

NOTE: parameter names are expected to be prefixed with a -
.PARAMETER Name
The name of the command parameter
.EXAMPLE
Lookup-ValidCommandParameterName -Name "-path"
Returns: -Path
.EXAMPLE
Lookup-ValidCommandParameterName -Name "-RECURSE"
Returns: -Recurse
.EXAMPLE
Lookup-ValidCommandParameterName -Name "-ParamNotFound"; Lookup-ValidCommandParameterName -Name "-paramNOTfound"
Returns: -ParamNotFound, -ParamNotFound
#>
function Lookup-ValidCommandParameterName {
  #region Function parameters
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [string]$Name
  )
  #endregion
  process {
    # look up name in lookup table and return
    # if not found (new function added within script), add to list and return
    if ($ValidCommandParameterNames.ContainsKey($Name)) {
      $ValidCommandParameterNames.Item($Name)
    } else {
      $script:ValidCommandParameterNames.Add($Name,$Name) > $null
      $Name
    }
  }
}
#endregion

#region Function: Lookup-ValidAttributeName

<#
.SYNOPSIS
Retrieves attribute name with correct casing
.DESCRIPTION
Retrieves the proper attribute name using the parameter attribute values
stored in the lookup table.
If Name is not found in the ValidAttributeNames lookup table, it is
added and returned as-is.  That means the first instance of the attribute 
that is encountered becomes the correct version, using its casing as the 
clean version.
.PARAMETER Name
The name of the attribute
.EXAMPLE
Lookup-ValidAttributeName -Name validatenotnull
Returns: ValidateNotNull
.EXAMPLE
Lookup-ValidAttributeName -Name AttribNotFound; Lookup-ValidAttributeName -Name ATTRIBNotFound
Returns: AttribNotFound, AttribNotFound
#>
function Lookup-ValidAttributeName {
  #region Function parameters
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [string]$Name
  )
  #endregion
  process {
    # look up name in lookup table and return
    # if not found, add to list and return
    if ($ValidAttributeNames.ContainsKey($Name)) {
      $ValidAttributeNames.Item($Name)
    } else {
      $script:ValidAttributeNames.Add($Name,$Name) > $null
      $Name
    }
  }
}
#endregion

#region Function: Lookup-ValidMemberName

<#
.SYNOPSIS
Retrieves member name with correct casing
.DESCRIPTION
Retrieves the proper member name using the member values stored in the lookup table.
If Name is not found in the ValidMemberNames lookup table, it is
added and returned as-is.  That means the first instance of the member 
that is encountered becomes the correct version, using its casing as the 
clean version.
.PARAMETER Name
The name of the member
.EXAMPLE
Lookup-ValidMemberName -Name valuefrompipeline
Returns: ValueFromPipeline
.EXAMPLE
Lookup-ValidMemberName -Name tostring
Returns: ToString
.EXAMPLE
Lookup-ValidMemberName -Name MemberNotFound; Lookup-ValidMemberName -Name MEMBERnotFound
Returns: MemberNotFound, MemberNotFound
#>
function Lookup-ValidMemberName {
  #region Function parameters
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [string]$Name
  )
  #endregion
  process {
    # look up name in lookup table and return
    # if not found, add to list and return
    if ($ValidMemberNames.ContainsKey($Name)) {
      $ValidMemberNames.Item($Name)
    } else {
      $script:ValidMemberNames.Add($Name,$Name) > $null
      $Name
    }
  }
}
#endregion

#region Function: Lookup-ValidVariableName

<#
.SYNOPSIS
Retrieves variable name with correct casing
.DESCRIPTION
Retrieves the proper variable name using the variable name values stored in the 
lookup table. If Name is not found in the ValidVariableNames lookup table, it is
added and returned as-is.  That means the first instance of the variable 
that is encountered becomes the correct version, using its casing as the 
clean version.
.PARAMETER Name
The name of the member
.EXAMPLE
Lookup-ValidVariableName -Name TRUE
Returns: true
.EXAMPLE
Lookup-ValidVariableName -Name MyVar; Lookup-ValidVariableName -Name MYVAR
Returns: MyVar, MyVar
#>
function Lookup-ValidVariableName {
  #region Function parameters
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [string]$Name
  )
  #endregion
  process {
    # look up name in lookup table and return
    # if not found, add to list and return
    if ($ValidVariableNames.ContainsKey($Name)) {
      $ValidVariableNames.Item($Name)
    } else {
      $script:ValidVariableNames.Add($Name,$Name) > $null
      $Name
    }
  }
}
#endregion
#endregion
#endregion

#region Add content to DestinationFileStreamWriter

<#
.SYNOPSIS
Copies a string into the destination stream.
.DESCRIPTION
Copies a string into the destination stream.
.PARAMETER Text
String to copy into destination stream.
#>
function Add-StringContentToDestinationFileStreamWriter {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [string]$Text
  )
  process {
    $script:DestinationStreamWriter.Write($Text)
  }
}

<#
.SYNOPSIS
Copies a section from the source string array into the destination stream.
.DESCRIPTION
Copies a section from the source string array into the destination stream.
.PARAMETER StartSourceIndex
Index in source string to start copy
.PARAMETER StartSourceLength
Length to copy
#>
function Copy-ArrayContentFromSourceArrayToDestinationFileStreamWriter {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [int]$StartSourceIndex,
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [int]$StartSourceLength
  )
  process {
    for ($i = $StartSourceIndex; $i -lt ($StartSourceIndex + $StartSourceLength); $i++) {
      $script:DestinationStreamWriter.Write($SourceScriptString[$i])
    }
  }
}
#endregion

#region Load content functions

<#
.SYNOPSIS
Reads content from source script into memory
.DESCRIPTION
Reads content from source script into memory
#>
function Import-ScriptContent {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    # get the file encoding of the file; it will be a type defined at System.Text.Encoding
    [System.Text.Encoding]$script:SourceFileEncoding = Get-DTWFileEncoding -Path $SourcePath

    # paths have already been validated so no testing of paths here
    # first load file as string array for use with Tokenize method
    $script:SourceScriptStringArray = [string[]](Get-Content -Path $SourcePath)
    # get error variable; works a bit differently in V3 CTP 2
    $ErrInfo = Get-Variable | Where-Object { $_.Name -eq "?" }
    if ($ErrInfo -ne $null -and $ErrInfo -eq $false) {
      Write-Error -Message "$($MyInvocation.MyCommand.Name):: error occurred getting content for SourceScriptStringArray with file: $SourcePath"
      return
    }
    # load file as a single String to access characters by original index
    $script:SourceScriptString = [System.IO.File]::ReadAllText($SourcePath)
    # get error variable; works a bit differently in V3 CTP 2
    $ErrInfo = Get-Variable | Where-Object { $_.Name -eq "?" }
    if ($ErrInfo -ne $null -and $ErrInfo -eq $false) {
      Write-Error -Message "$($MyInvocation.MyCommand.Name):: error occurred reading all text for getting content for SourceScriptString with file: $SourcePath"
      return
    }
  }
}
#endregion

#region Tokenize source script content

<#
.SYNOPSIS
Tokenizes code stored in $SourceScriptStringArray, stores in $SourceTokens
.DESCRIPTION
Tokenizes code stored in $SourceScriptStringArray, stores in $SourceTokens.  If an error
occurs, the error objects are written to the error stream.
#>
function Tokenize-SourceScriptContent {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    $Err = $null
    $script:SourceTokens = [System.Management.Automation.PSParser]::Tokenize($SourceScriptStringArray,[ref]$Err)
    if ($null -ne $Err -and $Err.Count) {
      Write-Error -Message "$($MyInvocation.MyCommand.Name):: error occurred tokenizing source content"
      $Err | ForEach-Object {
        Write-Error -Message "$($MyInvocation.MyCommand.Name):: $($_.Message) Content: $($_.Token.Content), line: $($_.Token.StartLine), column: $($_.Token.StartColumn)"
      }
      return
    }
  }
}
#endregion

#region Function to create destination file stream/writer

<#
.SYNOPSIS
Migrates content to destination stream.
.DESCRIPTION
Walks through tokens and for each copies (possibly modified) content to
destination stream.
#>
function Migrate-SourceContentToDestinationStream {
  #region Function parameters
  [CmdletBinding()]
  param()
  #endregion
  process {
    [int]$CurrentIndent = 0
    for ($i = 0; $i -lt $SourceTokens.Count; $i++) {
      # remove indents before writing GroupEnd
      if ($SourceTokens[$i].Type -eq 'GroupEnd') { $CurrentIndent -= 1 }

      #region Add indent to beginning of line
      # if last character was a NewLine or LineContinuation and current one isn't 
      # a NewLine nor a groupend, add indent prefix
      if ($i -gt 0 -and ($SourceTokens[$i - 1].Type -eq 'NewLine' -or $SourceTokens[$i - 1].Type -eq 'LineContinuation') `
           -and $SourceTokens[$i].Type -ne 'NewLine') {

        [int]$IndentToUse = $CurrentIndent
        # if last token was a LineContinuation, add an extra (one-time) indent
        # so indent lines continued (say cmdlet with many params)
        if ($SourceTokens[$i - 1].Type -eq 'LineContinuation') { $IndentToUse += 1 }
        # add the space prefix
        if ($IndentToUse -gt 0) {
          Add-StringContentToDestinationFileStreamWriter ($IndentText * $IndentToUse)
        }
      }
      #endregion

      # write the content of the token to the destination stream
      Write-TokenContentByType -SourceTokenIndex $i

      #region Add space after writing token
      if ($true -eq (Test-AddSpaceFollowingToken -TokenIndex $i)) {
        Add-StringContentToDestinationFileStreamWriter " "
      }
      #endregion

      #region Add indents after writing GroupStart
      if ($SourceTokens[$i].Type -eq 'GroupStart') { $CurrentIndent += 1 }
      #endregion
    }
  }
}
#endregion

#region Get new token functions

#region Function: Write-TokenContentByType

<#
.SYNOPSIS
Calls a token-type-specific function to writes token to destination stream.
.DESCRIPTION
This function calls a token-type-specific function to write the token's content
to the destination stream.  Based on the varied details about cleaning up 
the code based on type, expanding aliases, etc., this is best done in a function
for each type.  Even though many of these functions are similar (just 'write' 
content) to stream, let's keep these separate for easier maintenance.
The token's .Type is checked and the corresponding function Write-TokenContent_<type>
is called, which writes the token's content appropriately to the destination stream. 
There is a Write-TokenContent_* method for each entry on 
System.Management.Automation.PSTokenType
http://msdn.microsoft.com/en-us/library/system.management.automation.pstokentype(v=VS.85).aspx
.PARAMETER SourceTokenIndex
Index of current token in SourceTokens
#>
function Write-TokenContentByType {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [int]$SourceTokenIndex
  )
  process {
    # get name of function to call, based on type of source token
    [string]$FunctionName = "Write-TokenContent_" + $SourceTokens[$SourceTokenIndex].Type
    # call the appropriate new Token function, passing in source token, and return result
    & $FunctionName $SourceTokens[$i]
  }
}
#endregion

#region Write-TokenContent_* functions description
<#
Ok, normally I'd have help sections defined for each of these functions but they are
all incredibly similar and all that help is really just bloat.  (This file is getting 
bigger and bigger!).  Plus, these functions are private - not exported, so they will
never be accessed directly via Get-Help.

There are three important points to know about the Write-TokenContent_* functions:

1. There is a Write-TokenContent_* method for each Token.Type, that is, for each property 
value on the System.Management.Automation.PSTokenType enum.
See: http://msdn.microsoft.com/en-us/library/system.management.automation.pstokentype(v=VS.85).aspx

2. Each function writes the content to the destination stream using one of two ways:
  Add-StringContentToDestinationFileStreamWriter <string> 
    adds <string> to destination stream
  Copy-ArrayContentFromSourceArrayToDestinationFileStream
    copies the content directly from the source array to the destination stream
    
Why does the second function copy directly from the source array to the destination 
stream?  It has everything to do with escaped characters and whitespace issues.
Let's say your code has this line: Write-Host "Hello`tworld"
The problem is that if you store that "Hello`tworld" as a string, it becomes "Hello    world"
So, if you use the Token.Content value (whose value is a string), you have the expanded
string with the spaces.  This is fine if you are running the command and outputting the
results - they would be the same.  But if you are re-writing the source code, it's a
big problem.  By looking at the string "Hello    world" you don't know if that was its
original value or if "Hello`tworld" was.  You can easily re-escape the whitespace characters
within a string ("aa`tbb" -replace "`t","``t"), but you still don't know what the 
original was. I don't want to change your code incorrectly, I want it to be exactly the 
way that it was written before, just with correct whitespace, expanded aliases, etc. 
so storing the results in a stream is the way to go.  And as for the second function
Copy-ArrayContentFromSourceArrayToDestinationFileStream, copying each element from
the source array to the destination stream is the only way to keep the whitespace
intact.  If we extracted the content from the array into a string then wrote that to
the destination stream, we'd be back at square one.

This is important for these Token types: CommandArguments, String and Variable.
String is obvious.  Variable names can have whitespace using the { } notation; 
this is a perfectly valid, if insane, statement:  ${A`nB} = "hey now"
The Variable is named A`nB.
CommandArguments can also have whitespace AND will not have surrounding quotes.
For example, if this is tokenized: dir "c:\Program Files"
The "c:\Program Files" is tokenized as a String.  
However, the statement could be written as: dir c:\Program` Files
In this case, "c:\Program` Files" (no quotes!) is tokenized as a CommandArgument
that has whitespace in its value.  Joy.

3. Some functions will alter the value of the Token.Content before storing in the 
destination stream.  This is when the aliases are expanded, casing is fixed for 
command/parameter names, casing is fixed for types, keywords, etc.

Any special details will be described in each function.
#>
#endregion

#region Write token content for: Attribute
function Write-TokenContent_Attribute {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # check/replace Attribute value in ValidAttributeNames lookup table
    Add-StringContentToDestinationFileStreamWriter (Lookup-ValidAttributeName -Name $Token.Content)
  }
}
#endregion

#region Write token content for: Command
function Write-TokenContent_Command {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # check/replace CommandValue value in ValidCommandNames lookup table
    Add-StringContentToDestinationFileStreamWriter (Lookup-ValidCommandName -Name $Token.Content)
  }
}
#endregion

#region Write token content for: CommandArgument
function Write-TokenContent_CommandArgument {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # If you are creating a proxy function (function with same name as existing Cmdlet), it will be 
    # tokenized as a CommandArgument.  So, let's make sure the casing is the same by doing a lookup.
    # If the value is found in the valid list of CommandNames, it does not contain whitespace, so 
    # it should be safe to do a lookup and add the replacement text to the destination stream
    # otherwise copy the command argument text from source to destination.
    if ($ValidCommandNames.ContainsKey($Token.Content)) {
      Add-StringContentToDestinationFileStreamWriter (Lookup-ValidCommandName -Name $Token.Content)
    } else {
      # CommandArgument values can have whitespace, thanks to the escaped characters, i.e. dir c:\program` files
      # so we need to copy the value directly from the source to the destination stream.
      # By copying from the Token.Start with a length of Token.Length, we will also copy the 
      # backtick characters correctly
      Copy-ArrayContentFromSourceArrayToDestinationFileStreamWriter -StartSourceIndex $Token.Start -StartSourceLength $Token.Length
    }
  }
}
#endregion

#region Write token content for: CommandParameter
function Write-TokenContent_CommandParameter {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # check/replace CommandParameterName value in ValidCommandParameterNames lookup table
    Add-StringContentToDestinationFileStreamWriter (Lookup-ValidCommandParameterName -Name $Token.Content)
  }
}
#endregion

#region Write token content for: Comment
function Write-TokenContent_Comment {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add Comment Content as-is to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content
  }
}
#endregion

#region Write token content for: GroupEnd
function Write-TokenContent_GroupEnd {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add GroupEnd Content as-is to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content
  }
}
#endregion

#region Write token content for: GroupStart
function Write-TokenContent_GroupStart {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add GroupStart Content as-is to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content
  }
}
#endregion

#region Write token content for: KeyWord
function Write-TokenContent_KeyWord {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add KeyWord Content with lower case to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content.ToLower()
  }
}
#endregion

#region Write token content for: LoopLabel
function Write-TokenContent_LoopLabel {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # When tokenized, the loop label definition has a token type LoopLabel
    # and its content includes the colon prefix. However, when that loop label 
    # is used in a break statement, though, the loop label name is tokenized 
    # as a Member. So, in this function where the LoopLabel is defined, grab
    # the name (without the colon) and lookup (add if not found) to the 
    # Member lookup table.  When the loop label is used in the break statement,
    # it will look up in the Member table and use the same value, so the case
    # will be the same.

    # so, look up LoopLable name without colon in Members
    [string]$LookupNameInMembersNoColon = Lookup-ValidMemberName -Name ($Token.Content.Substring(1))
    # add to destination using lookup value but re-add colon prefix
    Add-StringContentToDestinationFileStreamWriter (":" + $LookupNameInMembersNoColon)
  }
}
#endregion

#region Write token content for: LineContinuation
function Write-TokenContent_LineContinuation {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add LineContinuation Content as-is to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content
  }
}
#endregion

#region Write token content for: Member
function Write-TokenContent_Member {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # check/replace Member value in ValidMemberNames lookup table
    Add-StringContentToDestinationFileStreamWriter (Lookup-ValidMemberName -Name $Token.Content)
  }
}
#endregion

#region Write token content for: NewLine
function Write-TokenContent_NewLine {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add NewLine using Windows standard to destination
    Add-StringContentToDestinationFileStreamWriter "`r`n"
  }
}
#endregion

#region Write token content for: Number
function Write-TokenContent_Number {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add Number Content as-is to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content
  }
}
#endregion

#region Write token content for: Operator
function Write-TokenContent_Operator {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add Operator Content with lower case to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content.ToLower()
  }
}
#endregion

#region Write token content for: Position
# I can't find much help info online about this type!  Just replicate content
function Write-TokenContent_Position {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    Add-StringContentToDestinationFileStreamWriter $Token.Content
  }
}
#endregion

#region Write token content for: StatementSeparator
function Write-TokenContent_StatementSeparator {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add StatementSeparator Content as-is to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content
  }
}
#endregion

#region Write token content for: String
function Write-TokenContent_String {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # String values can have whitespace, thanks to the escaped characters, i.e. "Hello`tworld"
    # so we need to copy the value directly from the source to the destination stream.
    # By copying from the Token.Start with a length of Token.Length, we will also copy the 
    # correct string boundary quote characters - even if it's a here-string. Nice!
    # Also, did you know that PowerShell supports multi-line strings that aren't here-strings?
    # This is valid:
    # $Message = "Hello
    # world"
    # It works.  I wish it didn't.
    Copy-ArrayContentFromSourceArrayToDestinationFileStreamWriter -StartSourceIndex $Token.Start -StartSourceLength $Token.Length
  }
}
#endregion

#region Write token content for: Type
function Write-TokenContent_Type {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # Let's try to get the correct case for types like [int] or [System.Exception]
    # If there isn't a . character in the type, it's a PowerShell type shortcut, so 
    #let's lowercase it.
    # If there is a . character in the type, let's try to create the type and get the 
    # fullname from the type.  If that fails (module/assembly not loaded) then just 
    # use the original value.
    $TypeName = $null
    if ($Token.Content.IndexOf(".") -eq -1) {
      # make sure content value is lower case
      $TypeName = $Token.Content.ToLower()
    } else {
      # need to wrap in try/catch in case type isn't loaded into memory, if so, just return
      try { $TypeName = ([type]::GetType($Token.Content,$true,$true)).FullName }
      catch { $TypeName = $Token.Content }
    }
    # if running in PowerShell v2, need to add [ ] around type name; PowerShell v3 is ok
    if ($Host.Version.Major -eq 2) { $TypeName = "[" + $TypeName + "]" }
    Add-StringContentToDestinationFileStreamWriter $TypeName
  }
}
#endregion

#region Write token content for: Variable
function Write-TokenContent_Variable {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # Variable names can have whitespace, thanks to the ${ } notation, i.e. ${A`nB} = 123
    # so we need to copy the value directly from the source to the destination stream.
    # By copying from the Token.Start with a length of Token.Length, we will also copy the 
    # variable markup, that is the $ or ${ }
    Copy-ArrayContentFromSourceArrayToDestinationFileStreamWriter -StartSourceIndex $Token.Start -StartSourceLength $Token.Length
  }
}
#endregion

#region Write token content for: Unknown
# I can't find much help info online about this type!  Just replicate content
function Write-TokenContent_Unknown {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $false)]
    [System.Management.Automation.PSToken]$Token
  )
  process {
    # add Unknown Content as-is to destination
    Add-StringContentToDestinationFileStreamWriter $Token.Content
  }
}
#endregion

#endregion

#region Generate destination file content

#region Function: Test-AddSpaceFollowingToken

<#
.SYNOPSIS
Returns $true if current token should be followed by a space, $false otherwise.
.DESCRIPTION
Returns $true if the current token, identified by the TokenIndex parameter, should
be followed by a space.  The logic that follows is basic: if a rule is found that 
determines a space should not be added $false is returned immediately.  If all rules
pass, $true is returned.  I normally do not like returning from within a function
in PowerShell but this logic is clean, the rules are well organized and it shaves some 
time off the process.
Here's an example: we don't want spaces between the [] characters for array index;
we want $MyArray[5], not $MyArray[ 5 ]. 
The rules will typically look at the current token BUT may want to check the next token 
as well.
.PARAMETER TokenIndex
Index of current token in $SourceTokens
#>
function Test-AddSpaceFollowingToken {
  #region Function parameters
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false,ValueFromPipeline = $false)]
    [int]$TokenIndex
  )
  #endregion
  process {
    # Notes: we sometimes need to check the next token to determine if we need to add a space.
    # In those checks, we need to make sure the current token isn't the last.

    # If a rule is found that space shouldn't be added, immediately return false.  If makes it
    # all the way through rules, return true.  To speed up this functioning, the rules that are
    # most likely to be useful are at the top.

    #region Don't write space after type NewLine
    if ($SourceTokens[$TokenIndex].Type -eq 'NewLine') { return $false }
    #endregion

    #region Don't write space after type Type if followed by GroupStart, Number, String, Type or Variable (for example [int]$Age or [int]"5")
    # don't write at space after, for example [int]
    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex].Type -eq 'Type' -and ('GroupStart','Number','String','Type','Variable') -contains $SourceTokens[$TokenIndex + 1].Type) { return $false }
    #endregion

    #region Don't write space if next token is StatementSeparator (;) or NewLine
    if (($TokenIndex + 1) -lt $SourceTokens.Count) {
      if ($SourceTokens[$TokenIndex + 1].Type -eq 'NewLine') { return $false }
      if ($SourceTokens[$TokenIndex + 1].Type -eq 'StatementSeparator') { return $false }
    }
    #endregion

    #region Don't add space before or after Operator [ or before Operator ]
    if ($SourceTokens[$TokenIndex].Type -eq 'Operator' -and $SourceTokens[$TokenIndex].Content -eq '[') { return $false }
    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex + 1].Type -eq 'Operator' -and ($SourceTokens[$TokenIndex + 1].Content -eq '[' -or $SourceTokens[$TokenIndex + 1].Content -eq ']')) { return $false }
    #endregion

    #region Don't write spaces before or after these Operators: . .. ::
    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex + 1].Type -eq 'Operator' -and $SourceTokens[$TokenIndex + 1].Content -eq '.') { return $false }
    if ($SourceTokens[$TokenIndex].Type -eq 'Operator' -and $SourceTokens[$TokenIndex].Content -eq '.') { return $false }

    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex + 1].Type -eq 'Operator' -and $SourceTokens[$TokenIndex + 1].Content -eq '..') { return $false }
    if ($SourceTokens[$TokenIndex].Type -eq 'Operator' -and $SourceTokens[$TokenIndex].Content -eq '..') { return $false }

    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex + 1].Type -eq 'Operator' -and $SourceTokens[$TokenIndex + 1].Content -eq '::') { return $false }
    if ($SourceTokens[$TokenIndex].Type -eq 'Operator' -and $SourceTokens[$TokenIndex].Content -eq '::') { return $false }
    #endregion

    #region Don't write space inside ( ) or $( ) groups
    if ($SourceTokens[$TokenIndex].Type -eq 'GroupStart' -and ($SourceTokens[$TokenIndex].Content -eq '(' -or $SourceTokens[$TokenIndex].Content -eq '$(')) { return $false }
    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex + 1].Type -eq 'GroupEnd' -and $SourceTokens[$TokenIndex + 1].Content -eq ')') { return $false }
    #endregion

    #region Don't write space if GroupStart ( { @{ followed by GroupEnd or NewLine
    if ($SourceTokens[$TokenIndex].Type -eq 'GroupStart' -and ($SourceTokens[$TokenIndex + 1].Type -eq 'GroupEnd' -or $SourceTokens[$TokenIndex + 1].Type -eq 'NewLine')) { return $false }
    #endregion

    #region Don't write space if writing Member or Attribute and next token is GroupStart
    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and 'Member','Attribute' -contains $SourceTokens[$TokenIndex].Type -and $SourceTokens[$TokenIndex + 1].Type -eq 'GroupStart') { return $false }
    #endregion

    #region Don't write space if writing Variable and next Operator token is [ (for example: $MyArray[3])
    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex].Type -eq 'Variable' -and $SourceTokens[$TokenIndex + 1].Type -eq 'Operator' -and $SourceTokens[$TokenIndex + 1].Content -eq '[') { return $false }
    #endregion

    #region Don't add space after Operators: , !
    if ($SourceTokens[$TokenIndex].Type -eq 'Operator' -and ($SourceTokens[$TokenIndex].Content -eq ',' -or $SourceTokens[$TokenIndex].Content -eq '!')) { return $false }
    #endregion

    #region Don't add space if next Operator token is: , ++ ;
    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex + 1].Type -eq 'Operator' -and
      ',','++',';' -contains $SourceTokens[$TokenIndex + 1].Content) { return $false }
    #endregion

    #region Don't add space after Operator > as in: 2>$null or 2>&1
    if ((($TokenIndex + 1) -lt $SourceTokens.Count) -and $SourceTokens[$TokenIndex].Type -eq 'Operator' -and $SourceTokens[$TokenIndex].Content -eq '2>' -and
      $SourceTokens[$TokenIndex + 1].Type -eq 'Variable' -and $SourceTokens[$TokenIndex + 1].Content -eq 'null') { return $false }
    if ($SourceTokens[$TokenIndex].Type -eq 'Operator' -and $SourceTokens[$TokenIndex].Content -eq '2>&1') { return $false }
    #endregion

    #region Don't add space after Keyword param
    if ($SourceTokens[$TokenIndex].Type -eq 'Keyword' -and $SourceTokens[$TokenIndex].Content -eq 'param') { return $false }
    #endregion

    #region Don't add space after CommandParameters with :<variable>
    # This is for switch params that are programmatically specified with a variable, such as:
    #   dir -Recurse:$CheckSubFolders
    if ($SourceTokens[$TokenIndex].Type -eq 'CommandParameter' -and $SourceTokens[$TokenIndex].Content[-1] -eq ':') { return $false }
    #endregion

    # return $true indicating add a space
    return $true
  }
}
#endregion
#endregion

#region Function: Edit-DTWCleanScript

<#
.SYNOPSIS
Cleans PowerShell script: indents code blocks, cleans/rearranges all 
whitespace, replaces aliases with commands, etc.
.DESCRIPTION
Cleans PowerShell script: indents code blocks, cleans/rearranges all 
whitespace, replaces aliases with commands, replaces parameter names with
proper casing, fixes case for [types], etc.  More specifically it:
 - properly indents code inside {}, [], () and $() groups
 - replaces aliases with the command names (dir -> Get-ChildItem)
 - fixes command name casing (get-childitem -> Get-ChildItem)
 - fixes parameter name casing (Test-Path -path -> Test-Path -Path)
 - fixes [type] casing
     changes all PowerShell shortcuts to lower ([STRING] -> [string])
     changes other types ([system.exception] -> [System.Exception]
       only works for types loaded into memory
 - cleans/rearranges all whitespace within a line
     many rules - see Test-AddSpaceFollowingToken to tweak

*****IMPORTANT NOTE: before running this on any script make sure you back
it up, commit any changes you have or run this on a copy of your script.
In the event this script screws up or the get encoding is incorrect (say
you have non-ANSI characters at the end of a long file with no BOM), I 
don't want your script to be damaged!

This version has been tested in PowerShell V2 and in V3 CTP 2.  Let me
know if you encounter any issues.

For all the alias and casing replacement rules above, the function works
on all items in memory, so it cleans your scripts using your own custom
function/parameter names and aliases as well.  The module caches all the
commands, aliases and parameter names when the module first loads.  If
you've added new commands to memory since loading the pretty printer 
module, you may want to reload it.

Also, this uses two spaces as an indent step.  To change this, edit the
line: [string]$script:IndentText = "  "

There are many rules when it comes to adding whitespace or not (see 
Test-AddSpaceFollowingToken).  Feel free to tweak this code and let me
know what you think is wrong or, at the very least, what should be 
configurable.  In version 2 I will expose configuration settings for
the items people feel strongly about.

This pretty printer module doesn't do everything - it's version 1.  
Version 2 (using PowerShell 3's new tokenizing/AST functionality) should
allow me to fix most of the deficiencies. But just so you know, here's
what it doesn't do:
 - change location of group openings, say ( or {, from same line to new
   line and vice-versa;
 - expand param names (Test-Path -Inc -> Test-Path -Include);
 - offer user options to control how space is reformatted.

See dtwconsulting.com or DansPowerShellStuff.blogspot.com for more 
information. And stay tuned for Pretty Printer version 2!

-Dan Ward dtward@gmail.com

.PARAMETER SourcePath
Path to the source PowerShell file
.PARAMETER DestinationPath
Path to write reformatted PowerShell.  If not specified rewrites file
in place.
.PARAMETER Quiet
If specified, does not output status text.
.EXAMPLE
Edit-DTWCleanScript -Source c:\P\S1.ps1 -Destination c:\P\S1_New.ps1
Gets content from c:\P\S1.ps1, cleans and writes to c:\P\S1_New.ps1
.EXAMPLE
Edit-DTWCleanScript -SourcePath c:\P\S1.ps1
Writes cleaned script results back into c:\P\S1.ps1
.EXAMPLE
dir c:\CodeFiles -Include *.ps1,*.psm1 -Recurse | Edit-DTWCleanScript
For each .ps1 and .psm1 file, cleans and rewrites back into same file
#>
function Edit-DTWCleanScript {
  #region Function parameters
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
    [Alias("FullName")]
    [string]$SourcePath,
    [string]$DestinationPath,
    [switch]$Quiet
  )
  #endregion
  process {
    # initialize all script-level variables used in a cleaning process
    Initialize-ProcessVariables

    # if not initialize properly, exit
    if ($false -eq $ModuleInitialized) {
      Write-Error -Message "$($MyInvocation.MyCommand.Name):: Module not properly initialized; exiting."
      return
    }

    #region Parameter validation
    #region SourcePath must exist; if not, exit
    if ($false -eq (Test-Path -Path $SourcePath)) {
      Write-Error -Message "$($MyInvocation.MyCommand.Name) :: SourcePath does not exist: $SourcePath"
      return
    }
    #endregion

    # resolve path name so make sure we have full name
    $SourcePath = Resolve-Path -Path $SourcePath

    #region Source file must contain content
    if ((Get-Item -Path $SourcePath).Length -eq 0 -or ([System.IO.File]::ReadAllText($SourcePath)).Trim() -eq "") {
      Write-Host "File contains no content: $SourcePath"
      return
    }
    #endregion
    #endregion

    [datetime]$StartTime = Get-Date

    #region Set source, destination and temp file paths
    $script:PathSource = $SourcePath
    # if no destination passed, use source
    if ($DestinationPath -eq "") {
      $script:PathDestination = $PathSource
    } else {
      $script:PathDestination = $DestinationPath
    }
    $script:PathDestinationTemp = $PathDestination + ".pspp"
    #endregion

    #region Read source script content into memory
    if (!$Quiet) { Write-Host "Reading source: $PathSource" }
    $Err = $null
    Import-ScriptContent -EV $Err
    # if an error occurred importing content, it will be written from Import-ScriptContent
    # so in this case just exit
    if ($null -ne $Err) { return }
    #endregion

    #region Tokenize source script content
    $Err = $null
    if (!$Quiet) { Write-Host "Tokenizing script content" }
    Tokenize-SourceScriptContent -EV Err
    # if an error occurred tokenizing content, it will be written from Tokenize-SourceScriptContent
    # so in this case just initialize the process variables and exit
    if ($null -ne $Err -and $Err.Count) {
      Initialize-ProcessVariables
      return
    }
    #endregion

    # create stream writer in try/catch so can dispose if error
    try {
      #region
      # Destination content is stored in a stream which is written to the file at the end.
      # It has to be a stream; storing it in a string or string builder loses the original
      # values (with regard to escaped characters).

      # if parent destination folder doesn't exist, create
      [string]$Folder = Split-Path -Path $PathDestinationTemp -Parent
      if ($false -eq (Test-Path -Path $Folder)) { New-Item -Path $Folder -ItemType Directory > $null }

      # create file stream writer, overwrite - don't append, and use same encoding as source file
      $script:DestinationStreamWriter = New-Object System.IO.StreamWriter $PathDestinationTemp,$false,$SourceFileEncoding
      #endregion
      # create new tokens for destination script content
      if (!$Quiet) { Write-Host "Migrate source content to destination format" }
      Migrate-SourceContentToDestinationStream
    } catch {
      Write-Error -Message "$($MyInvocation.MyCommand.Name) :: error occurred during processing"
      Write-Error -Message "$($_.ToString())"
      return
    } finally {
      #region Flush and close file stream writer
      if (!$Quiet) { Write-Host "Write destination file: $PathDestination" }
      if ($null -ne $script:DestinationStreamWriter) {
        $script:DestinationStreamWriter.Flush()
        $script:DestinationStreamWriter.Close()
        $script:DestinationStreamWriter.Dispose()
        #region Replace destination file with destination temp (which has updated content)
        # if destination file already exists, remove it
        if ($true -eq (Test-Path -Path $PathDestination)) { Remove-Item -Path $PathDestination -Force }
        # rename destination temp to destination
        Rename-Item -Path $PathDestinationTemp -NewName (Split-Path -Path $PathDestination -Leaf)
        #endregion
      }
      #endregion
      if (!$Quiet) { Write-Host ("Finished in {0:0.000} seconds.`n" -f ((Get-Date) - $StartTime).TotalSeconds) }
    }
  }
}
Export-ModuleMember -Function Edit-DTWCleanScript
#endregion

# 'main'; initialize the module, initial load of lookup table values, etc.
Initialize-Module

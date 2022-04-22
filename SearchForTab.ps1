$filePath = $args[0]
$extens =  "*.tab"
#$extensstring = $args[1] 
#$extens =  $extensstring.Split(',')

$sequences = Get-ChildItem -Force -Recurse $filePath -ErrorAction SilentlyContinue -Include $extens | Where-Object { ($_.PSIsContainer -eq $false) } | Select-Object FullName, Name, DirectoryName

$regexstring = '.*KEY\s*NUMBER\(\s*(?<digits>\d*).*\r\n'

$regex = new-object System.Text.RegularExpressions.Regex ($regexstring, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

$outputarray= @()
ForEach ($file  in $sequences )
{
    $contents = Get-Content $file.FullName -Raw 
    if ($contents -ne $null)
    {
        $matches = $regex.Matches($contents, 0)
        if ($matches.Count -gt 0)
        {
            foreach($match in $matches)
            {
                $s = $match.Groups[1].Value

                $outputprops = 
                @{
                    Digits = $s
                    Line = (new-object System.String( $match.Groups[0].Value)).Trim()
                    Name=$file.Name
                    Path=$file.DirectoryName
                    FullPath = $file.FullName
                }    
                $outputobject = New-Object PSObject –Property $outputprops
                $outputarray +=  $outputobject
            }
        }
    }
}

#$outputarray |   Format-Table -AutoSize -Wrap Sequencer, Name, Path, FullName  
$outputarray | Sort-Object Line | Format-Table -AutoSize -Wrap Line, Digits, Name  
Write-Host "Number of Keys = " $outputarray.Length

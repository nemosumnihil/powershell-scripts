$filePath = $args[0]
$extens =  "*.cs"
#$extensstring = $args[1] 
#$extens =  $extensstring.Split(',')
$sequences = Get-ChildItem -Force -Recurse $filePath -ErrorAction SilentlyContinue -Include $extens | Where-Object { ($_.PSIsContainer -eq $false) } | Select-Object FullName, Name, DirectoryName

$regexstring = 'new Sequence\s*\(\s*\"(?<sequencer>.*)\"\s*,\s*(?<cachenum>\d*)\s*\)'

$regex = new-object System.Text.RegularExpressions.Regex ($regexstring, [System.Text.RegularExpressions.RegexOptions]::Compiled)


$outputarray= @()
ForEach ($file  in $sequences )
{
    $contents = Get-Content $file.FullName    
    if ($contents -ne $null)
    {
        $match = $regex.Match($contents)
        if ($match.Success)
        {
            $s = $match.Groups[1].Value
            $s = $s.Substring(0,[System.Math]::Min(18, $s.Length))
            if ($s -ne "seq_dbtable132_ESN")
            {
                #For some reason seq_dbtable132_ESN gives me a problem where it returns the whole match group
                $s = $match.Groups[1].Value;
            }

            $outputprops = 
            @{
                Sequencer = $s
                CacheNum = $match.Groups[2].Value
                Name=$file.Name
                Path=$file.DirectoryName
                FullPath = $file.FullName
            }    
            $outputobject = New-Object PSObject –Property $outputprops
            $outputarray +=  $outputobject
        }
    }
}

#$outputarray |   Format-Table -AutoSize -Wrap Sequencer, Name, Path, FullName  
$outputarray |   Format-Table -AutoSize -Wrap Sequencer, CacheNum, FullPath  
Write-Host "Number of Sequencers = " $outputarray.Length

$filePath = $args[0]
$extens =  "*.seq"
#$extensstring = $args[1] 
#$extens =  $extensstring.Split(',')
$sequences = Get-ChildItem -Force -Recurse $filePath -ErrorAction SilentlyContinue -Include $extens | Where-Object { ($_.PSIsContainer -eq $false) } | Select-Object FullName, Name, DirectoryName

$regexstring = '\s*(?:create sequence|CREATE SEQUENCE)\s(?<sequencename>.*)\s*minvalue\s(?<minvalue>\d*)\s*maxvalue\s(?<maxvalue>\d*)\s*start\swith\s(?<startwith>\d*)\s*increment\sby\s(?<incrementby>\d*)\s*cache\s(?<cache>\d*);'
$regexstring2 = '\s*CREATE SEQUENCE\s(?<sequencename>.*)\s*START\sWITH\s(?<startwith>\d*)\s*INCREMENT\sBY\s(?<incrementby>\d*)\s*MINVALUE\s(?<minvalue>\d*)\s*CACHE\s(?<cache>\d*)'
$regexstring3 = '\s*CREATE SEQUENCE\s(?<sequencename>.*)\s*START\sWITH\s(?<startwith>\d*)\s*MAXVALUE\s(?<maxvalue>\d*)\s*MINVALUE\s(?<minvalue>\d*)\s*NOCYCLE\s*CACHE\s(?<cache>\d*)'

$regex = new-object System.Text.RegularExpressions.Regex ($regexstring, [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
$regex2 = new-object System.Text.RegularExpressions.Regex ($regexstring2, [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
$regex3 = new-object System.Text.RegularExpressions.Regex ($regexstring3, [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)


$outputarray= @()
ForEach ($file  in $sequences )
{
    $contents = Get-Content $file.FullName    
    $ismatch = $regex.IsMatch($contents)
    if ($ismatch)
    {
        $regexmatch = $regex.Match($contents)
        $outputprops = 
        @{
            SeqName=$regexmatch.Groups[1].Value
            MinValue=$regexmatch.Groups[2].Value
            MaxValue=$regexmatch.Groups[3].Value
            Digits=$regexmatch.Groups[3].Value.Length
            StartsWith=$regexmatch.Groups[4].Value
            IncrementBy=$regexmatch.Groups[5].Value
            Cache=$regexmatch.Groups[6].Value
            Name=$file.Name
            Path=$file.DirectoryName
        }    
        $outputobject = New-Object PSObject –Property $outputprops
        $outputarray +=  $outputobject
        
    }
    else
    {
        if ($regex2.IsMatch($contents))
        {
            $regexmatch2 = $regex2.Match($contents)
            $outputprops2 = 
            @{
                SeqName=$regexmatch2.Groups[1].Value
                StartsWith=$regexmatch2.Groups[2].Value
                IncrementBy=$regexmatch2.Groups[3].Value
                MinValue=$regexmatch2.Groups[4].Value
                Cache=$regexmatch2.Groups[5].Value

                MaxValue='-'
                Digits='-'
                Name=$file.Name
                Path=$file.DirectoryName
         
            }    
            $outputobject2 = New-Object PSObject –Property $outputprops2
            $outputarray +=  $outputobject2
        }
        else
        {
            if ($regex3.IsMatch($contents))
            {
                $regexmatch3 = $regex3.Match($contents)
                $outputprops3 = 
                @{
                    SeqName=$regexmatch3.Groups[1].Value
                    StartsWith=$regexmatch3.Groups[2].Value
                    MaxValue=$regexmatch3.Groups[3].Value
                    MinValue=$regexmatch3.Groups[4].Value
                    Cache=$regexmatch3.Groups[5].Value
                    Digits=$regexmatch3.Groups[3].Value.Length
                    IncrementBy='-'
                    Name=$file.Name
                    Path=$file.DirectoryName
                }    
                $outputobject3 = New-Object PSObject –Property $outputprops3
                $outputarray +=  $outputobject3
            }
            else
            {
                Write-Host $contents -BackgroundColor Red
            }
        }
    }
}


#$sequences |   Format-Table -AutoSize *  
$outputarray |   Format-Table -AutoSize -Wrap SeqName, MinValue, MaxValue, Digits, StartsWith, IncrementBy, Cache, Name, Path  
Write-Host "Number of Sequences = " $outputarray.Length

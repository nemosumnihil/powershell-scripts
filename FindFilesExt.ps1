$filePath = $args[0]
$extensstring = $args[1] 
$extens =  $extensstring.Split(',')
$sequences = Get-ChildItem -Force -Recurse $filePath -ErrorAction SilentlyContinue -Include $extens | Where-Object { ($_.PSIsContainer -eq $false) } | Select-Object Name,Directory
$sequences | Format-Table -AutoSize *  
Write-Host "Number of Sequences = " + $sequences.Length
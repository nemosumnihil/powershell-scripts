"`n" 
write-Host "---------------------------------------------" -ForegroundColor Yellow 
$filePath = Read-Host "Please Enter File Path to Search" 
write-Host "---------------------------------------------" -ForegroundColor Green 
$fileName = Read-Host "Please Enter File Name to Search" 
write-Host "---------------------------------------------" -ForegroundColor Yellow 
"`n" 
 
Get-ChildItem -Recurse -Force $filePath -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and  ( $_.Name -like "*$fileName*") } | Select-Object Name,Directory| Format-Table -AutoSize * 
 
write-Host "------------END of Result--------------------" -ForegroundColor Magenta 
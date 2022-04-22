Get-Content $args[0] |
  Select-String 'Project\(' |
    ForEach-Object {
      $projectParts = $_ -Split '[,=]' | ForEach-Object { $_.Trim('[ "{}]') };
      New-Object PSObject -Property @{
        File = $projectParts[2];
        Name = $projectParts[1]
      }
    }
function Write-CustomProgress {
    param(
        [int]$PercentComplete
    )

    $barLength = 50
    $filledLength = [Math]::Round(($PercentComplete / 100) * $barLength)
    $bar = ('#' * $filledLength).PadRight($barLength, '-')
    Write-Host "`r[$bar] `t$PercentComplete%" -NoNewline 
    
}

function Split-ZipFile {

    $dateadd = $false
    Write-host "`n"

    [string]$dirpath =    $(Write-host    "~? Pfad zum Dateienordner       " -ForegroundColor DarkYellow -NoNewline; Read-Host)

    if($dirpath -like '*"*') {
        $dirpath = $dirpath.Trim('"')
    }

    if((Test-Path $dirpath) -eq $false)
    {
        Write-host "Bitte einen Gültigen Pfad eingeben" -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        return
    }


    [string]$destpath =   $(Write-host   "~? Pfad zum Ausgabeordner       " -ForegroundColor DarkYellow -NoNewline; Read-Host) 

    
    if($destpath -like '*"*') {
        $destpath = $destpath.Trim('"')
    }

    
    if((Test-Path $destpath) -eq $false)
    {
        Write-host "Bitte einen Gültigen Pfad eingeben" -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        return
    }

    [int]$targetSizeKB =  $(Write-host  "~? Grösse in KB                 " -ForegroundColor DarkYellow -NoNewline; Read-Host) 

    if($targetSizeKB -isnot [int])
    {
        Write-host "`n"
        Write-Host "Bitte eine gerade Zahl eingeben" -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        return
    }

    [string]$nameconv =   $(Write-host   "~? Namenskonvension der Zips    " -ForegroundColor DarkYellow -NoNewline; Read-Host)
    $dateaddinput =       $(Write-host   "~? Datum hinzufügen (n)         " -ForegroundColor DarkYellow -NoNewline; Read-Host) 
    
    if($dateaddinput -ne 'n' -or $dateaddinput -ne 'N')
    {
        [bool]$dateadd = $true
    }


    [string]$csvheaders = $(Write-host "~? CSV Headers                  " -ForegroundColor DarkYellow -NoNewline; Read-Host)

    if($csvheaders -notlike "*,*")
    {
        Write-host "`n"
        Write-Host "CSV headers mit kommas trennen" -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        return
    }


    <#
        Patterntesting
    #>


    cls

    Write-host "
               ___ __             _           
   _________  / (_) /_     ____  (_)___  _____
  / ___/ __ \/ / / __/____/_  / / / __ \/ ___/
 (__  ) /_/ / / / /_/_____// /_/ / /_/ (__  ) 
/____/ .___/_/_/\__/      /___/_/ .___/____(_)
    /_/                        /_/            
" -ForegroundColor Yellow


    Write-host "`n"
    $(Write-host "~? Welche Teile der Dateinamen sollte das CSV beinhalten" -ForegroundColor DarkYellow) 
    Write-host "`n"

    $testfile = Get-ChildItem -Path $dirpath -File | Select-Object -First 1
    $extension = $testfile.Extension
    
    [string]$testfilename = $testfile.Name
    
    $matchesarray = @()
    $patterntomatch = @()
    $patternid = @()

    if($testfilename -match '[-_]?(\d+)[-_]')
    {
        if($matchesarray -notcontains $Matches[1])
        {   
            $matchesarray += $Matches[1]
            
            Write-host $Matches[1] -NoNewline
            $spaces = 30 - $Matches[1].Length

            if($spaces -lt 0)
            {
                $spaces = 0
            }

            Write-Host (" " * $spaces) -NoNewline

            $option = Read-Host
    
            if($option -ne 'n' -or $option -ne 'n')
            {
                $patternid += $Matches[1]
                $patterntomatch += '[-_]?(\d+)[-_]'
            }
        }
    }
    
    #check if fn match String at Start
    
    if($testfilename -match '[-_]?(\D+)[-_]')
    {
        if($matchesarray -notcontains $Matches[1])
        {
            $matchesarray += $Matches[1]
            
            Write-host $Matches[1] -NoNewline
            $spaces = 30 - $Matches[1].Length

            if($spaces -lt 0)
            {
                $spaces = 0
            }

            Write-Host (" " * $spaces) -NoNewline

            $option = Read-Host
    
            if($option -ne 'n' -or $option -ne 'n')
            {
                $patternid += $Matches[1]
                $patterntomatch += '[-_]?(\D+)[-_]'
            }
        }
    }
    
    
    #check if fn match Digits after and before [-_]
    
    if($testfilename -match '[-_](\d+)[-_]')
    {
        if($matchesarray -notcontains $Matches[1])
        {
            $matchesarray += $Matches[1]
            
            Write-host $Matches[1] -NoNewline
            $spaces = 30 - $Matches[1].Length

            if($spaces -lt 0)
            {
                $spaces = 0
            }

            Write-Host (" " * $spaces) -NoNewline

            $option = Read-Host
    
           if($option -ne 'n' -or $option -ne 'n')
            {
                $patternid += $Matches[1]
                $patterntomatch += '[-_](\d+)[-_]'
            }
        }
    }
    
    #check if fn match String after and before [-_]
    
    if($testfilename -match '[-_](\D+)[-_]')
    {
        if($matchesarray -notcontains $Matches[1])
        {
            $matchesarray += $Matches[1]
            
            Write-host $Matches[1] -NoNewline
            $spaces = 30 - $Matches[1].Length

            if($spaces -lt 0)
            {
                $spaces = 0
            }

            Write-Host (" " * $spaces) -NoNewline

            $option = Read-Host
    
            if($option -ne 'n' -or $option -ne 'n')
            {
                $patternid += $Matches[1]
                $patterntomatch += '[-_](\D+)[-_]'
            }
        }
    }
    
    #check if fn match digit before extension
    
    if($testfilename -match ('(\d+)' + $extension))
    {
        if($matchesarray -notcontains $Matches[1])
        {
            $matchesarray += $Matches[1]
            
            Write-host $Matches[1] -NoNewline
            $spaces = 30 - $Matches[1].Length

            if($spaces -lt 0)
            {
                $spaces = 0
            }

            Write-Host (" " * $spaces) -NoNewline

            $option = Read-Host
    
            if($option -ne 'n' -or $option -ne 'n')
            {
                $patternid += $Matches[1]
                $patterntomatch += ('(\d+)' + $extension)
            }
        }
    }
    
    #check if fn match String before extension
    
    if($testfilename -match ('(\D+)' + $extension))
    {
        if($matchesarray -notcontains $Matches[1])
        {
            $matchesarray += $Matches[1]
            
            Write-host $Matches[1] -NoNewline
            $spaces = 30 - $Matches[1].Length

            if($spaces -lt 0)
            {
                $spaces = 0
            }

            Write-Host (" " * $spaces) -NoNewline

            $option = Read-Host
    
            if($option -ne 'n' -or $option -ne 'n')
            {
                $patternid += $Matches[1]
                $patterntomatch += ('(\D+)' + $extension)
            }
        }
    }
  
     $matchesarray += $testfilename
     Write-host "Filename" -NoNewline
     $spaces = 30 - "Filename".Length

     if($spaces -lt 0)
     {
         $spaces = 0
     }

     Write-Host (" " * $spaces) -NoNewline

     $option = Read-Host


    if($option -ne 'n' -or $option -ne 'n')
    {
        $patternid += $testfilename
        $patterntomatch += 'wholefilename'
    }


     Write-host "Date" -NoNewline
     $spaces = 30 - "Date".Length

     if($spaces -lt 0)
     {
         $spaces = 0
     }

     Write-Host (" " * $spaces) -NoNewline

     $option = Read-Host

     if($option -ne 'n' -or $option -ne 'n')
     {
          do{
        
            Write-host "Zeitformat" -NoNewline
            $spaces = 30 - "zeitformat".Length

            if($spaces -lt 0)
            {
                $spaces = 0
            }

            Write-Host (" " * $spaces) -NoNewline

            [string]$format = Read-Host

            if($format -match '[yMdhms]')
            {
               $success = $true
               $date = Get-Date -Format $format
            }
            else
            {
                  $success = $false
            }

          }while($success -eq $false)

            


         $patternid += $date
         $patterntomatch += 'date-' + $format
     }



    $csvheadersarray = $csvheaders -split ','
 

    $neworder = @()

    for($i = 0; $i -lt $csvheadersarray.Length; $i++)
    {
        cls

        Write-host "
               ___ __             _           
   _________  / (_) /_     ____  (_)___  _____
  / ___/ __ \/ / / __/____/_  / / / __ \/ ___/
 (__  ) /_/ / / / /_/_____// /_/ / /_/ (__  ) 
/____/ .___/_/_/\__/      /___/_/ .___/____(_)
    /_/                        /_/            
" -ForegroundColor Yellow
        
        Write-host "`n"
        Write-host $csvheadersarray[$i]'/' -ForegroundColor DarkYellow
        Write-host "`n"

        foreach($pattern in $patternid)
        {

            $index = $patternid.IndexOf($pattern)

             
            Write-host $pattern -NoNewline
            $spaces = 40 - $pattern.Length

            if($spaces -lt 0)
            {
                $spaces = 0
            }

            Write-Host (" " * $spaces) -NoNewline

            Write-host $index

        }

        Write-host "`n`n"
        Write-host "~?" -NoNewline -ForegroundColor DarkYellow
        $spaces = 40 - "~?".Length
        Write-Host (" " * $spaces) -NoNewline
        $setheader = Read-host



        if($setheader -like "*-*")
        {
            $numbers = $setheader -replace '-', ''
            $temp = @()

            for($a = 0; $a -lt $numbers.Length; $a++)
            {
                [int]$currentnumber = $numbers[$a].ToString()
                $temp += $patterntomatch[$currentnumber]
                
            }
            $neworder += $temp -Join '|'
        }

        else
        {
            $neworder += $patterntomatch[$setheader]
        }

    }
    
    cls

    Write-host "
               ___ __             _           
   _________  / (_) /_     ____  (_)___  _____
  / ___/ __ \/ / / __/____/_  / / / __ \/ ___/
 (__  ) /_/ / / / /_/_____// /_/ / /_/ (__  ) 
/____/ .___/_/_/\__/      /___/_/ .___/____(_)
    /_/                        /_/            
" -ForegroundColor Yellow

    Get-ChildItem -Path $dirpath -File | Compress-Archive -DestinationPath ($destpath + $nameconv) 

    [string]$zipFilePath = $destpath + $nameconv + ".zip"

   
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($zipFilePath)
    
    $files = @($zipArchive.Entries | Where-Object { $_.CompressedLength -lt $targetSizeKB * 1KB })

    $splitIndex = 1
    $currentSize = 0
    $currentFiles = New-Object System.Collections.Generic.List[System.IO.Compression.ZipArchiveEntry]
    $todaysdate = Get-Date -Format "yyyyMMddhhmmss"


    foreach ($file in $files) {

    Write-CustomProgress ((($files.IndexOf($file) + 1) / $files.Count) * 100)

        if($file.Fullname -ne "data.csv"){

        $currentSize += $file.CompressedLength

        if ($currentSize -gt $targetSizeKB * 1KB) {

            if($dateadd -eq $true)
            {
                $newZipPath = [System.IO.Path]::ChangeExtension($zipFilePath, "_${splitIndex}_${todaysdate}.zip")
            }
            else
            {
                $newZipPath = [System.IO.Path]::ChangeExtension($zipFilePath, "_${splitIndex}.zip")
            }
            
            $newZipPath = [System.IO.Path]::ChangeExtension($zipFilePath, "_${splitIndex}_${todaysdate}.zip")
            CreateNewZipWithCSV -newZipPath $newZipPath -currentFiles $currentFiles $csvheaders $neworder
            $splitIndex++
            $currentSize = $file.CompressedLength
            $currentFiles.Clear()
        }

        $currentFiles.Add($file)
        }
    }

    if ($currentFiles.Count -gt 0) {

        if($dateadd -eq $true)
        {
            $newZipPath = [System.IO.Path]::ChangeExtension($zipFilePath, "_${splitIndex}_${todaysdate}.zip")
        }
        else
        {
            $newZipPath = [System.IO.Path]::ChangeExtension($zipFilePath, "_${splitIndex}.zip")
        }

        CreateNewZipWithCSV -newZipPath $newZipPath -currentFiles $currentFiles $csvheaders $neworder
    }

    $zipArchive.Dispose()


    $zipFiles = Get-ChildItem -Path $destpath -Filter "$nameconv*.zip"

    foreach($zip in $zipFiles)
    {
        $newFileName = ($zip.BaseName -replace "\.", "") + $zip.Extension
        $newFilePath = Join-Path -Path $destpath -ChildPath $newFileName

         if ($newFileName -ne $zip.Name) {
            Rename-Item -Path $zip.FullName -NewName $newFilePath
        }
    }


    Remove-Item -Path $zipFilePath



    Write-host "`n"
    Write-host "Prozess fertig" -ForegroundColor Green
    Write-host "`n"

    Start-Sleep -Seconds 3

}

function CreateNewZipWithCSV {
    param (
        [string]$newZipPath,
        [System.Collections.Generic.List[System.IO.Compression.ZipArchiveEntry]]$currentFiles,
        $csvheaders,
        [array]$patterns
    )

    $csvContent = "$csvheaders`r`n"

    foreach ($entry in $currentFiles) {

        $fileName = $entry.Name
        $matcharray = @()

        foreach($pattern in $patterns)
        {
            if($pattern -like '*|*')
            {

                [array]$splitpattern = $pattern -split '\|'


                $temp = @()

                for($i = 0; $i -lt $splitpattern.Count; $i++)
                {
                    if($splitpattern[$i] -eq 'wholefilename')
                    {
                        $temp += $fileName
                    }
                    
                    elseif($splitpattern[$i] -like "*date-*")
                    {
                        $substring = ($splitpattern[$i]).substring(5)
                        $temp += Get-Date -Format $substring
                    }

                    elseif ($fileName -match $splitpattern[$i])
                    {
                        $temp += $Matches[1]
                    }
                    
                }

                $temp = $temp -join '-'


                $matcharray += $temp 

            }

            else
            {

                if($pattern -eq 'wholefilename')
                {
                    $matcharray += $fileName
                }

                elseif($pattern -like "*date-*")
                {
                    $substring = $pattern.substring(5)
                    $matcharray += Get-Date -Format $substring
                }
                
                elseif ($fileName -match $pattern)
                {
                    $matcharray += $Matches[1]
                }

            }

        }

         $csvcont = $matcharray -join ','

         $csvContent += " $csvcont`r`n"
        
    }

    $newZipArchive = [System.IO.Compression.ZipFile]::Open($newZipPath, [System.IO.Compression.ZipArchiveMode]::Create)
    foreach ($entry in $currentFiles) {


        $entryStream = $entry.Open()
        $newEntry = $newZipArchive.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
        $newEntryStream = $newEntry.Open()
        $entryStream.CopyTo($newEntryStream)
        $newEntryStream.Close()
        $entryStream.Close()
    }

    # Add CSV to the ZIP
    $csvEntry = $newZipArchive.CreateEntry("data.csv", [System.IO.Compression.CompressionLevel]::Optimal)
    $csvStream = $csvEntry.Open()
    $writer = New-Object System.IO.StreamWriter($csvStream)
    $writer.Write($csvContent)
    $writer.Close()


    $newZipArchive.Dispose()
}

do{


try{
    cls

    Write-host "
               ___ __             _           
   _________  / (_) /_     ____  (_)___  _____
  / ___/ __ \/ / / __/____/_  / / / __ \/ ___/
 (__  ) /_/ / / / /_/_____// /_/ / /_/ (__  ) 
/____/ .___/_/_/\__/      /___/_/ .___/____(_)
    /_/                        /_/            
" -ForegroundColor Yellow

    Split-ZipFile
}
catch {
    Write-host "`n"
    Write-host "Unerwarterter Fehler" -ForegroundColor Red
    $_
    Write-host "`n"

}

}while($true)

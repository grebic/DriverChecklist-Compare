cd "P:\ANG_System_Files"

function Load-Dll
{
    param(
        [string]$assembly
    )
    Write-Host "Loading $assembly"

    $driver = $assembly
    $fileStream = ([System.IO.FileInfo] (Get-Item $driver)).OpenRead();
    $assemblyBytes = new-object byte[] $fileStream.Length
    $fileStream.Read($assemblyBytes, 0, $fileStream.Length) | Out-Null;
    $fileStream.Close();
    $assemblyLoaded = [System.Reflection.Assembly]::Load($assemblyBytes);
}

function Get-ComparisonObjects
{
    param([Smartsheet.Api.Models.Sheet]$sheet)

    Write-Host "Getting Sheet $($sheet.Name) Comparison Objects"

    $data = $sheet.Rows | foreach {
        
        [pscustomobject]@{
            Attachments = $_.Attachments;
            RowId = $_.Id;
            RowNumber = $_.RowNumber;
            PoCol = $_.Cells[0].ColumnId;
            Po = $_.Cells[0].Value;
            Parent = $_.ParentId;
            SKUCol = $_.Cells[18].ColumnId;
            SKU = $_.Cells[18].Value;

        }                                                  
    }| where {![string]::IsNullOrWhiteSpace($_.Po)}  

    Write-Host "$($data.Count) Returned"      
    return $data                                           
}   

function Get-DriverComparisonObjects
{
    param([Smartsheet.Api.Models.Sheet]$sheet)

    Write-Host "Getting Sheet $($sheet.Name) Comparison Objects"

    $data = $sheet.Rows | foreach {
        
        $archiveCheckVal = $false

        if($_.Cells[11].Value -eq $true)
        {
            $archiveCheckVal = $true
        }
        [pscustomobject]@{
            Attachments = $_.Attachments;
            RowId = $_.Id;
            RowNumber = $_.RowNumber;
            Parent = $_.ParentId;
            PoNumCol = $_.Cells[9].ColumnId;
            PoNum = $_.Cells[9].Value;
            ModifiedCol = $_.Cells[10].ColumnId;
            Modified = $_.Cells[10].Value;
            ArchiveCol = $_.Cells[11].ColumnId;
            Archive = $archiveCheckVal;
        }                                                  
    } | where {![string]::IsNullOrWhiteSpace($_.PoNum)}   

    Write-Host "$($data.Count) Returned"      
    return $data                                            
}   

function Find-MatchingSku
{
    foreach ($DriveCO in $DriveCOs)
    {
        $skuFound = $false

        $driveQR = "$($DriveCO.PoNum)"
        Out-File -FilePath "P:\ANG_System_Files\commonFormsUsedInScripts\QRcheck\driveQR.txt" -InputObject $driveQR

        if ($DriveCO.Archive -eq $false)
        {
            foreach ($ptCO in $ptCOs)
            {
                $ptQR = "$($ptCO.SKU)"

                Out-File -FilePath "P:\ANG_System_Files\commonFormsUsedInScripts\QRcheck\ptQR.txt" -InputObject $ptQR

                $DriveQRPath = "P:\ANG_System_Files\commonFormsUsedInScripts\QRcheck\driveQR.txt"
                $PtQRPath = "P:\ANG_System_Files\commonFormsUsedInScripts\QRcheck\ptQR.txt"
                
                $QRdrive = (Get-FileHash $DriveQRPath).hash 
                $QRpt = (Get-FileHash $PtQRPath).hash
               
                if ($QRdrive -eq $QRpt)
                {
                    $skuFound = $true
                    break
                }
            }
                
            if ($skuFound)
            { 
                if (![string]::IsNullOrWhiteSpace($DriveCO.PoNum) -and ![string]::IsNullOrWhiteSpace($ptCO.SKU))
                {
                    Write-Host "SKU found.  Updating data on PT."
                    
                    $destinationCell = [Smartsheet.Api.Models.Cell]::new()
                    $destinationCell.ColumnId  = $ptDestinationCol.Id
                    $destinationCell.Value     = "ITEM SCANNED TO ANG TRUCK"

                    $shippedCell = [Smartsheet.Api.Models.Cell]::new()
                    $shippedCell.ColumnId  = $ptShippedCol.Id
                    $shippedCell.Value     = if ($DriveCO.Modified -ne $null){$DriveCO.Modified} else {[string]::Empty}

                    $row = [Smartsheet.Api.Models.Row]::new()
                    $row.Id = $ptCO.RowId
                    $row.Cells = [Smartsheet.Api.Models.Cell[]]@($destinationCell, $shippedCell)

                    $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
                }
            }
        }
    }
}

Write-Host "Loading Dlls"
Load-Dll ".\smartsheet-csharp-sdk.dll"                     
Load-Dll ".\RestSharp.dll"
Load-Dll ".\Newtonsoft.Json.dll"
Load-Dll ".\NLog.dll"

while($true)
{
    Write-Host "Fab Log to Driver List to Product Tracker system starting up."

    $DriveId = "" 
    $ptId    = ""
    
    $token      = ""
    $smartsheet = [Smartsheet.Api.SmartSheetBuilder]::new()
    $builder    = $smartsheet.SetAccessToken($token)
    $client     = $builder.Build()
    $includes   =  @([Smartsheet.Api.Models.SheetLevelInclusion]::ATTACHMENTS)
    $includes   = [System.Collections.Generic.List[Smartsheet.Api.Models.SheetLevelInclusion]]$includes
    
    $Drive  = $client.SheetResources.GetSheet($DriveId, $includes, $null, $null, $null, $null, $null, $null);
    $pt     = $client.SheetResources.GetSheet($ptId, $includes, $null, $null, $null, $null, $null, $null);
    
    $DriveCOs  = Get-DriverComparisonObjects $Drive
    $ptCOs     = Get-ComparisonObjects $pt
    
    $ptDestinationCol = $pt.Columns | where {$_.Title -eq ("Destination")}
    $ptShippedCol     = $pt.Columns | where {$_.Title -eq ("Shipped Date")}
   
   Find-MatchingSku 

   Write-Host "Going to sleep."
   
   Start-Sleep -Seconds 10
}

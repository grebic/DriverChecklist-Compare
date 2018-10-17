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
            
            RowId = $_.Id;
            DestinationCol = $_.Cells[11].ColumnId;
            Destination = $_.Cells[11].Value;
            ModifiedCol = $_.Cells[16].ColumnId;
            Modified = $_.Cells[16].Value;
            QrCodeCol = $_.Cells[18].ColumnId;
            QrCode = $_.Cells[18].Value;
        }                                                  
    }| where {![string]::IsNullOrWhiteSpace($_.QrCode)}  

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
            
            RowId = $_.Id;
            DestinationCol = $_.Cells[7].ColumnId;
            Destination = $_.Cells[7].Value;
            QrCodeCol = $_.Cells[9].ColumnId;
            QrCode = $_.Cells[9].Value;
            ModifiedCol = $_.Cells[10].ColumnId;
            Modified = $_.Cells[10].Value;
        }                                                  
    } | where {![string]::IsNullOrWhiteSpace($_.QrCode)}   

    Write-Host "$($data.Count) Returned"      
    return $data                                            
}   

function Merge-ComparisonObjectsWithProductTracker
{
    param([PSCustomObject[]]$COs, [string]$comment)

    foreach($CO in $COs)
    {
        $matches = $pts | where {$_.QrCode -eq $CO.QrCode}

        if($matches)
        {
            foreach ($match in $matches)
            {
                if ($($match.Destination) -ne "ITEM SCANNED TO ANG TRUCK FOR FIELD")
                {
                    Write-Host ""
                    Write-Host $CO.QrCode

                    $ptDestinationCol = $ptSheet.Columns | where {$_.Title -eq ("Destination")}
                    $ptShippedCol     = $ptSheet.Columns | where {$_.Title -eq ("Shipped Date")}

                    $destinationCell = [Smartsheet.Api.Models.Cell]::new()
                    $destinationCell.ColumnId  = $ptDestinationCol.Id
                    $destinationCell.Value     = $comment
                    
                    $shippedCell = [Smartsheet.Api.Models.Cell]::new()
                    $shippedCell.ColumnId  = $ptShippedCol.Id
                    $shippedCell.Value     = if ([string]::IsNullOrWhiteSpace($match.Modified)){Get-Date} else{$match.Modified}
                    
                    $row = [Smartsheet.Api.Models.Row]::new()
                    $row.Id = $match.RowId
                    $row.Cells = [Smartsheet.Api.Models.Cell[]]@($destinationCell, $shippedCell)
                    
                    $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
                }
            }
        }
    }
}

function Update-WarehouseInventory
{
    param([PSCustomObject[]]$COs, [Smartsheet.Api.Models.Sheet]$sheet, [long]$sheetId)

    foreach($CO in $COs)
    {
        $matches = $pts | where {$_.QrCode -eq $CO.QrCode}

        if($matches)
        {
            foreach ($match in $matches)
            {
                if ([string]::IsNullOrWhiteSpace($CO.Destination))
                {
                    if ($($match.Destination) -eq "ITEM SCANNED TO ANG TRUCK FOR FIELD")
                    {
                        Write-Host ""
                        Write-Host "$($CO.QrCode) `nTransferred out of $($sheet.Name)"

                        $DestinationCol = $Sheet.Columns | where {$_.Title -eq ("Destination")}
                        $ShippedCol     = $Sheet.Columns | where {$_.Title -eq ("Shipped Date")}

                        $destinationCell = [Smartsheet.Api.Models.Cell]::new()
                        $destinationCell.ColumnId  = $DestinationCol.Id
                        $destinationCell.Value     = "ITEM SCANNED TO ANG TRUCK FOR FIELD"
                        
                        $shippedCell = [Smartsheet.Api.Models.Cell]::new()
                        $shippedCell.ColumnId  = $ShippedCol.Id
                        $shippedCell.Value     = if ([string]::IsNullOrWhiteSpace($co.Modified)){Get-Date} else{$co.Modified}
                        
                        $row = [Smartsheet.Api.Models.Row]::new()
                        $row.Id = $co.RowId
                        $row.Cells = [Smartsheet.Api.Models.Cell[]]@($destinationCell, $shippedCell)
                        
                        $updateRow = $client.SheetResources.RowResources.UpdateRows($sheetId, [Smartsheet.Api.Models.Row[]]@($row))
                    }
                }
            }
        }
    }
}

while($true)
{

    Write-Host "Loading Dlls"
    Load-Dll ".\smartsheet-csharp-sdk.dll"                     
    Load-Dll ".\RestSharp.dll"
    Load-Dll ".\Newtonsoft.Json.dll"
    Load-Dll ".\NLog.dll"
    
    $DriveId    = "" 
    $ptId       = ""
    $4042Id     = ""
    $puyallupId = ""
    
    $token      = ""
    $smartsheet = [Smartsheet.Api.SmartSheetBuilder]::new()
    $builder    = $smartsheet.SetAccessToken($token)
    $client     = $builder.Build()
    $includes   =  @([Smartsheet.Api.Models.SheetLevelInclusion]::ATTACHMENTS)
    $includes   = [System.Collections.Generic.List[Smartsheet.Api.Models.SheetLevelInclusion]]$includes
    
    $driveSheet    = $client.SheetResources.GetSheet($DriveId, $includes, $null, $null, $null, $null, $null, $null);
    $ptSheet       = $client.SheetResources.GetSheet($ptId, $includes, $null, $null, $null, $null, $null, $null);
    $4042Sheet     = $client.SheetResources.GetSheet($4042Id, $includes, $null, $null, $null, $null, $null, $null);
    $puyallupSheet = $client.SheetResources.GetSheet($puyallupId, $includes, $null, $null, $null, $null, $null, $null);

    $ptDestinationCol = $pt.Columns | where {$_.Title -eq ("Destination")}
    $ptShippedCol     = $pt.Columns | where {$_.Title -eq ("Shipped Date")}

    $driveCOs    = Get-DriverComparisonObjects $driveSheet
    $ptCOs       = Get-ComparisonObjects $ptSheet
    $4042COs     = Get-ComparisonObjects $4042Sheet
    $puyallupCOs = Get-ComparisonObjects $puyallupSheet

    Write-Host "Resuming"
    Write-Host ""

    $pts       = $ptCOs| where {$_.QrCode -like "*All New Glass" } 
    $drives    = $driveCOs | where {$_.QrCode -like "*All New Glass" } 
    $4042s     = $4042COs| where {$_.QrCode -like "*All New Glass" } 
    $puyallups = $puyallupCOs | where {$_.QrCode -like "*All New Glass" }  
    
    Write-Host ""
    Write-Host "Checking 4042 Warehouse" 

    Merge-ComparisonObjectsWithProductTracker -COs $4042s -comment "ITEM SCANNED TO 4042 WAREHOUSE" 
    Update-WarehouseInventory -COs $4042s -sheet $4042Sheet -sheetId $4042Id

    Write-Host ""
    Write-Host "Checking Puyallup Warehouse"

    Merge-ComparisonObjectsWithProductTracker -COs $puyallups -comment "ITEM SCANNED TO PUYALLUP WAREHOUSE" 
    Update-WarehouseInventory -COs $puyallups -sheet $puyallupSheet -sheetId $puyallupId

    Write-Host ""
    Write-Host "Checking Driver List"

    Merge-ComparisonObjectsWithProductTracker -COs $drives -comment "ITEM SCANNED TO ANG TRUCK FOR FIELD" 
        
    Write-Host ""
    Write-Host "Pausing..."
    Start-Sleep -Seconds 20
}

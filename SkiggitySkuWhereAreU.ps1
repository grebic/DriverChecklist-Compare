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
            ShippedCol = $_.Cells[10].ColumnId;
            Shipped = $_.Cells[10].Value;
            DestinationCol = $_.Cells[11].ColumnId;
            Destination = $_.Cells[11].Value;
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

Write-Host "Loading Dlls"
Load-Dll ".\smartsheet-csharp-sdk.dll"                     
Load-Dll ".\RestSharp.dll"
Load-Dll ".\Newtonsoft.Json.dll"
Load-Dll ".\NLog.dll"

$DriveId = "" 
$ptId    = ""

$token      = ""
$smartsheet = [Smartsheet.Api.SmartSheetBuilder]::new()
$builder    = $smartsheet.SetAccessToken($token)
$client     = $builder.Build()
$includes   =  @([Smartsheet.Api.Models.SheetLevelInclusion]::ATTACHMENTS)
$includes   = [System.Collections.Generic.List[Smartsheet.Api.Models.SheetLevelInclusion]]$includes

$driveSheet  = $client.SheetResources.GetSheet($DriveId, $includes, $null, $null, $null, $null, $null, $null);
$ptSheet     = $client.SheetResources.GetSheet($ptId, $includes, $null, $null, $null, $null, $null, $null);
    
while($true)
{
    $driveCOs  = Get-DriverComparisonObjects $driveSheet
    $ptCOs     = Get-ComparisonObjects $ptSheet

    Write-Host "Resuming"
    Write-Host ""

    $drives = $driveCOs | where {$_.PoNum -ne $null }    
    $pts = $ptCOs| where {$_.Sku -ne $null } 
    
    foreach($pt in $pts) 
    {
        $matches = $drives | where {$_.PoNum -eq $pt.SKU} 
        if($matches)
        {
            foreach ($match in $matches)
            {
                Write-Host $pt.SKU
                write-host ""

                $ptDestinationCol = $ptSheet.Columns | where {$_.Title -eq ("Destination")}
                $ptShippedCol     = $ptSheet.Columns | where {$_.Title -eq ("Shipped Date")}

                $destinationCell = [Smartsheet.Api.Models.Cell]::new()
                $destinationCell.ColumnId  = $ptDestinationCol.Id
                $destinationCell.Value     = "ITEM SCANNED TO ANG TRUCK"
                
                $shippedCell = [Smartsheet.Api.Models.Cell]::new()
                $shippedCell.ColumnId  = $ptShippedCol.Id
                $shippedCell.Value     = if ([string]::IsNullOrWhiteSpace($pt.Shipped)){Get-Date} else{$pt.Shipped}
                
                $row = [Smartsheet.Api.Models.Row]::new()
                $row.Id = $pt.RowId
                $row.Cells = [Smartsheet.Api.Models.Cell[]]@($destinationCell, $shippedCell)
                
                $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
            }
        }
    }
        
    Write-Host "Pausing..."
    Start-Sleep -Seconds 10
}

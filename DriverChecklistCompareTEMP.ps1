# by GrEcHkO

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
        $checkVal = $false
        $trackedCheckVal = $false
        $finishedCheckVal = $false

        if($_.Cells[3].Value -eq $true)
        {
            $checkVal = $true
        }
        
        if($_.Cells[16].Value -eq $true)
        {
            $trackedCheckVal = $true
        } 

        if($_.Cells[12].Value -eq $true)
        {
            $finishedCheckVal = $true
        }
        [pscustomobject]@{
            Attachments = $_.Attachments;
            RowId = $_.Id;
            RowNumber = $_.RowNumber;
            Parent = $_.ParentId;
            PoCol = $_.Cells[0].ColumnId;
            Po = $_.Cells[0].Value;
            JobsCol = $_.Cells[1].ColumnId;
            Jobs = $_.Cells[1].Value;
            DescCol = $_.Cells[2].ColumnId;
            Desc = $_.Cells[2].Value;
            CheckCol = $_.Cells[3].ColumnId;
            Check = $checkVal;
            SupplierCol = $_.Cells[4].ColumnId;
            Supplier = $_.Cells[4].Value;
            AssignCol = $_.Cells[5].ColumnId;
            Assign = $_.Cells[5].Value;
           
            DestinationCol = $_.Cells[11].ColumnId;
            Destination = $_.Cells[11].Value;
            FinishedCol = $_.Cells[12].ColumnId;
            Finished = $finishedCheckVal;
            DueCol = $_.Cells[14].ColumnId;
            Due = $_.Cells[14].Value;
            DeliveryCol = $_.Cells[15].ColumnId;
            Delivery = $_.Cells[15].Value;
            TrackedCol = $_.Cells[16].ColumnId;
            Tracked = $trackedCheckVal;
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
        $checkVal = $false
        $trackedCheckVal = $false
        $finishedCheckVal = $false
        $archiveCheckVal = $false

        if($_.Cells[2].Value -eq $true)
        {
            $checkVal = $true
        }
        
        if($_.Cells[8].Value -eq $true)
        {
            $trackedCheckVal = $true
        } 

        if($_.Cells[3].Value -eq $true)
        {
            $finishedCheckVal = $true
        }

        if($_.Cells[11].Value -eq $true)
        {
            $archiveCheckVal = $true
        }
        [pscustomobject]@{
            Attachments = $_.Attachments;
            RowId = $_.Id;
            RowNumber = $_.RowNumber;
            Parent = $_.ParentId;
            DayCol = $_.Cells[0].ColumnId;
            Day = $_.Cells[0].Value;
            DueCol = $_.Cells[1].ColumnId;
            Due = $_.Cells[1].Value;
            CompletedCol = $_.Cells[2].ColumnId; #########hidden
            Completed = $checkVal;
            CheckCol = $_.Cells[3].ColumnId;
            Check = $finishedCheckVal;
            SupplierCol = $_.Cells[4].ColumnId;#########hidden
            Supplier = $_.Cells[4].Value;
            AssignCol = $_.Cells[5].ColumnId;#########hidden
            Assign = $_.Cells[5].Value;
            JobNameCol = $_.Cells[6].ColumnId;
            JobName = $_.Cells[6].Value;
            MainCol = $_.Cells[7].ColumnId;
            Main = $_.Cells[7].Value;
            TrackedCol = $_.Cells[8].ColumnId;
            Tracked = $trackedCheckVal;
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

function Get-AttachmentFromSmartsheet
{
    param (
        [long]$attachmentId,
        [long]$sheetId
    )
    
    Write-Host "Getting Attachement $attachmentId of Sheet $sheetId"

    
    try
    {
        $attachment = $client.SheetResources.AttachmentResources.GetAttachment($sheetId,$attachmentId)
    }
    catch
    {
        Write-Error $_.Exception.Message
        Write-Host ""
    }
    
    $downloads = New-Item -ItemType Directory ".\downloads" -Force 

    $filepath = "$($downloads.Fullname)\$($attachment.Name)"

    Write-Host "Downloading $filepath"

    Invoke-WebRequest -Uri $attachment.Url -OutFile $filepath 
    
    Get-Item $filepath
}
 
function Save-AttachmentToSheetRow
{
    param(
        [long]$sheetId,
        [long]$rowId,
        [System.IO.FileInfo]$file,
        [string]$mimeType
    )

    Write-Host "Saving $($file.Fullname) to Sheet $sheetId"

    $result = $client.SheetResources.RowResources.AttachmentResources.AttachFile($sheetId, $rowId, $file.FullName, $mimeType)

    return $result
}
     

function Merge-DriverChecklistWithProductTracker
{
     param(
        [pscustomobject[]]$orbitalRecords,
        [long]$orbitalId
    )

    foreach ($orbitalRecord in $orbitalRecords)
    {
        $descFound = $false
        $skuFound = $false

        if ($orbitalRecord.Archive -eq $false)
        {
            foreach ($ptCO in $ptCOs)
            {
                $orbitalQR = "$($orbitalRecord.PoNum)"
                $ptQR = "$($ptCO.SKU)"
                $noParent = $false

                if ($orbitalQR -eq $ptQR)
                {
                    $descFound = $true
                    break
                }

                if ($orbitalQR -eq $ptQR)
                {
                    $skuFound = $true
                    break
                }
            }
                
            if ($skuFound)
            { 
                if (![string]::IsNullOrWhiteSpace($orbitalRecord.PoNum) -and ![string]::IsNullOrWhiteSpace($ptCO.SKU))
                {
                    Write-Host "SKU found.  Updating data on PT."
                    
                    $destinationCell = [Smartsheet.Api.Models.Cell]::new()
                    $destinationCell.ColumnId  = $ptDestinationCol.Id
                    $destinationCell.Value     = "ITEM SCANNED TO ANG TRUCK"

                    $shippedCell = [Smartsheet.Api.Models.Cell]::new()
                    $shippedCell.ColumnId  = $ptShippedCol.Id
                    $shippedCell.Value     = if ($orbitalRecord.Modified -ne $null){$orbitalRecord.Modified} else {[string]::Empty}

                    $row = [Smartsheet.Api.Models.Row]::new()
                    $row.Id = $ptCO.RowId
                    $row.Cells = [Smartsheet.Api.Models.Cell[]]@($destinationCell, $shippedCell)

                    $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
                }
            }
            
            elseif ($orbitalRecord.PoNum -eq $ptCO.Po)
            {
                if ($descFound)
                {
                     if ([string]::IsNullOrWhiteSpace($ptCO.Parent))
                     {
                         $noParent = $true
                         
                     }
            
                    if ($noParent)
                    {
                        Write-Host "Updating Product Tracker with Driver Checklist $($orbitalRecord.PoNum) $($orbitalRecord.Main)"
            
                        $jobsCell = [Smartsheet.Api.Models.Cell]::new()
                        $jobsCell.ColumnId   = $ptJobsCol.Id
                        $JobsCell.Value      =  if ($orbitalRecord.JobName -ne $null){$orbitalRecord.JobName} else {[string]::Empty}
                        
                        $descCell = [Smartsheet.Api.Models.Cell]::new()
                        $descCell.ColumnId   = $ptDescCol.Id
                        $descCell.Value      =  if ($orbitalRecord.Main -ne $null){$orbitalRecord.Main} else {[string]::Empty}
            
                        $checkCell = [Smartsheet.Api.Models.Cell]::new()
                        $checkCell.ColumnId  = $ptCheckCol.Id
                        $checkCell.Value     = $orbitalRecord.Completed
            
                        $assignCell = [Smartsheet.Api.Models.Cell]::new()
                        $assignCell.COlumnId    = $ptAssignCol.Id
                        $assignCell.Value       = if ($orbitalRecord.Assign -ne $null){$orbitalRecord.Assign} else {"ianz@allnewglass.com"}
                        
                        $supCell = [Smartsheet.Api.Models.Cell]::new()
                        $supCell.COlumnId    = $ptSupplierCol.Id
                        $supCell.Value       = if (($orbitalRecord.Completed -eq $false) -and ($orbitalRecord.Supplier -eq "In Shop FAB")){$suppliers["$fabId"]} else {$suppliers["$DriveId"]}
            
                        $finishedCell = [Smartsheet.Api.Models.Cell]::new()
                        $finishedCell.ColumnId  = $ptFinishedCol.Id
                        $finishedCell.Value     = $orbitalRecord.Check
            
                        $row = [Smartsheet.Api.Models.Row]::new()
                        $row.Id = $ptCO.RowId
                        $row.Cells = [Smartsheet.Api.Models.Cell[]]@($JobsCell, $checkCell, $supCell, $assignCell, $finishedCell)
                        try
                        {
                            $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
            
                            $reference  = if ($orbitalRecord.Attachments.Name -ne $null){$orbitalRecord.Attachments.Name} else {[string]::Empty}
                            $difference = if ($ptCO.Attachments.Name -ne $null){$ptCO.Attachments.Name} else {[string]::Empty}
            
                            $compareResults = Compare-Object -ReferenceObject $reference -DifferenceObject $difference -IncludeEqual -SyncWindow ([int]::MaxValue)
            
                            $missing = $compareResults | where SideIndicator -eq '<='
                            $missingAttachments = $missing.InputObject
                            $attachmentsToGet = $orbitalRecord.Attachments | where Name -in $missingAttachments
            
                            foreach($attachment in $attachmentsToGet)
                            {
                                Write-Host "Adding missing attachment $($attachment.name)"
                                $file = Get-AttachmentFromSmartsheet -attachmentId $attachment.Id -sheetId $orbitalId
                                $result = Save-AttachmentToSheetRow -sheetId $ptId -rowId $newRow.Id -file $file.FullName -mimeType $attachment.MimeType
                            }
                        }
                        catch
                        {
                            Write-Error $_.Exception.Message
                            Write-Host ""
                        }
            
                        if ($orbitalRecord.Check -eq $true)
                        {
                            $dateCell = [Smartsheet.Api.Models.Cell]::new()
                            $dateCell.ColumnId  = $ptShippedCol.Id
                            $dateCell.Value     = $orbitalRecord.Modified
            
                            $row = [Smartsheet.Api.Models.Row]::new()
                            $row.Id = $ptCO.RowId
                            $row.Cells = [Smartsheet.Api.Models.Cell[]]@($dateCell)
                            try
                            {
                                $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
                            }
                            catch
                            {
                                Write-Error $_.Exception.Message
                                Write-Host ""
                            }
                        }
            
                        if (![string]::IsNullOrWhiteSpace($orbitalRecord.Due))
                        {
                            $dateCell = [Smartsheet.Api.Models.Cell]::new()
                            $dateCell.ColumnId  = $ptAnticipatedCol.Id
                            $dateCell.Value     = $orbitalRecord.Due
            
                            $row = [Smartsheet.Api.Models.Row]::new()
                            $row.Id = $ptCO.RowId
                            $row.Cells = [Smartsheet.Api.Models.Cell[]]@($dateCell)
                            try
                            {
                                $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
                            }
                            catch
                            {
                                Write-Error $_.Exception.Message
                                Write-Host ""
                            }
                        }
                    }
                }
            }
            
            else
            {
                if (!($descFound -or $skuFound))
                {
                    if (![string]::IsNullOrWhiteSpace($orbitalRecord.PoNum))
                    {
                        if ($orbitalRecord.Tracked -eq "$true")
                        {
                            Write-Host "Adding to Product Tracker from Driver Checklist $($orbitalRecord.Po) $($orbitalRecord.Main)"
                
                            $poCell = [Smartsheet.Api.Models.Cell]::new()
                            $poCell.ColumnId     = $ptPoCol.Id
                            $poCell.Value        = if ($orbitalRecord.PoNum -ne $null){$orbitalRecord.PoNum} else {[string]::Empty}
                
                            $jobsCell = [Smartsheet.Api.Models.Cell]::new()
                            $jobsCell.ColumnId   = $ptJobsCol.Id
                            $JobsCell.Value      =  if ($orbitalRecord.JobName -ne $null){$orbitalRecord.JobName} else {[string]::Empty}
                
                            $descCell = [Smartsheet.Api.Models.Cell]::new()
                            $descCell.ColumnId   = $ptDescCol.Id
                            $descCell.Value      =  if ($orbitalRecord.Main -ne $null){$orbitalRecord.Main} else {[string]::Empty}
                
                            $checkCell = [Smartsheet.Api.Models.Cell]::new()
                            $checkCell.ColumnId  = $ptCheckCol.Id
                            $checkCell.Value     = $orbitalRecord.Completed
                
                            $supCell = [Smartsheet.Api.Models.Cell]::new()
                            $supCell.COlumnId    = $ptSupplierCol.Id
                            $supCell.Value       = if (($orbitalRecord.Completed -eq $false) -and ($orbitalRecord.Supplier -eq "In Shop FAB")){$suppliers["$fabId"]} else {$suppliers["$DriveId"]}
                
                            $assignCell = [Smartsheet.Api.Models.Cell]::new()
                            $assignCell.COlumnId    = $ptAssignCol.Id
                            $assignCell.Value       = if ($orbitalRecord.Assign -ne $null){$orbitalRecord.Assign} else {"ianz@allnewglass.com"}
                
                            $finishedCell = [Smartsheet.Api.Models.Cell]::new()
                            $finishedCell.ColumnId  = $ptFinishedCol.Id
                            $finishedCell.Value     = $orbitalRecord.Check
                
                            $row = [Smartsheet.Api.Models.Row]::new()
                            $row.ToBottom = $true
                            $row.Cells = [Smartsheet.Api.Models.Cell[]]@($poCell,$jobsCell,$descCell,$checkCell,$supCell, $assignCell, $finishedCell) 
                 
                            try
                            {
                                $newRow = $client.SheetResources.RowResources.AddRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
                
                                foreach($attachment in $orbitalRecord.Attachments)
                                {
                                    $file = Get-AttachmentFromSmartsheet -attachmentId $attachment.Id -sheetId $orbitalId
                                    $result = Save-AttachmentToSheetRow -sheetId $ptId -rowId $newRow.Id -file $file.FullName -mimeType $attachment.MimeType
                                }
                            }
                            catch
                            {
                                Write-Error $_.Exception.Message
                                Write-Host ""
                            }
                            
                        }
                    }
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

    $DriveId = "470920482580356" 
    $fabId   = "6217793554147204"
    $ptId    = "5779331080316804"
    
    $suppliers = @{
        $fabId   = "In Shop FAB";
        $DriveId = "DRIVER CHECKLIST";
    }
    
    $token      = ""
    $smartsheet = [Smartsheet.Api.SmartSheetBuilder]::new()
    $builder    = $smartsheet.SetAccessToken($token)
    $client     = $builder.Build()
    $includes   =  @([Smartsheet.Api.Models.SheetLevelInclusion]::ATTACHMENTS)
    $includes   = [System.Collections.Generic.List[Smartsheet.Api.Models.SheetLevelInclusion]]$includes
    
    Write-Host "Loading Sheets"
    
    $Drive  = $client.SheetResources.GetSheet($DriveId, $includes, $null, $null, $null, $null, $null, $null);
    $fab    = $client.SheetResources.GetSheet($fabId, $includes, $null, $null, $null, $null, $null, $null);
    $pt     = $client.SheetResources.GetSheet($ptId, $includes, $null, $null, $null, $null, $null, $null);
    
    Write-Host "Comparing Objects"
    
    $DriveCOs  = Get-DriverComparisonObjects $Drive
    $fabCOs    = Get-ComparisonObjects $fab
    $ptCOs     = Get-ComparisonObjects $pt
    
    Write-Host "Identifying Driver Checklist Columns"
    
    $DriveDayCol       = $Drive.Columns | where {$_.Title -eq ("Day of Week")}
    $DriveDueCol       = $Drive.Columns | where {$_.Title -eq ("Due Date")}
    $DriveCompletedCol = $Drive.Columns | where {$_.Title -eq ("Completed")}
    $DriveCheckCol     = $Drive.Columns | where {$_.Title -eq ("Check")}
    $DriveVendorCol    = $Drive.Columns | where {$_.Title -eq ("Vendor")}
    $DriveAssignCol    = $Drive.Columns | where {$_.Title -eq ("Assigned To")}
    $DriveJobCol       = $Drive.Columns | where {$_.Title -eq ("Job Name")}
    $DriveMainCol      = $Drive.Columns | where {$_.Title -eq ("Main")}
    $DriveTrackingCol  = $Drive.Columns | where {$_.Title -eq ("Tracking")}
    $DrivePurchaseCol  = $Drive.Columns | where {$_.Title -eq ("PO / SKU")}
    $DriveModifiedCol  = $Drive.Columns | where {$_.Title -eq ("Modified")}
    
    $ptPoCol          = $pt.Columns | where {$_.Title -eq ("PO/WO #")}
    $ptJobsCol        = $pt.Columns | where {$_.Title -eq ("Job")}
    $ptDescCol        = $pt.Columns | where {$_.Title -eq ("Description")}
    $ptCheckCol       = $pt.Columns | where {$_.Title -eq ("Completed")}
    $ptSupplierCol    = $pt.Columns | where {$_.Title -eq ("Vendor")}
    $ptAssignCol      = $pt.Columns | where {$_.Title -eq ("Assigned To")}
    $ptDestinationCol = $pt.Columns | where {$_.Title -eq ("Destination")}
    $ptFinishedCol    = $pt.Columns | where {$_.Title -eq ("Finished")}
    $ptDueCol         = $pt.Columns | where {$_.Title -eq ("Due Date")}
    $ptDeliveryCol    = $pt.Columns | where {$_.Title -eq ("Delivery Method")}
    $ptTrackedCol     = $pt.Columns | where {$_.Title -eq ("Tracked")}
    $ptShippedCol     = $pt.Columns | where {$_.Title -eq ("Shipped Date")}
    $ptAnticipatedCol = $pt.Columns | where {$_.Title -eq ("Anticipated Received")}
   
   Merge-DriverChecklistWithProductTracker -orbitalRecords $DriveCOs -orbitalId $DriveId 

   Write-Host "Driver Checklist to Product Tracker finished."
  
   Write-Host "Driver List to Product Tracker taking a nap......"
   
   Start-Sleep -Seconds 10

}
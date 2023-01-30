# Read me 1/29/2023
  # This script will convert AADDS's LDAP 1644 DGrep output into Excel pivot tables for workload analysis, to use this script:
  #    1. Run DGrep with 
        # source
        # | sort by TIMESTAMP asc
        # | project-rename LDAPServer=RoleInstance, TimeGenerated=PreciseTimeStamp, StartingNode=Data1, Filter=Data2, VisitedEntries=Data3, ReturnedEntries=Data4, Client=Data5, SearchScope=Data6, AttributeSelection=Data7, ServerControls=Data8, UsedIndexes=Data9, PagesReferenced=Data10, PagesReadFromDisk=Data11, PagesPreReadFromDisk=Data12, CleanPagesModified=Data13, DirtyPagesModified=Data14, SearchTimeMS=Data15, AttributesPreventingOptimization=Data16, User=Data17
        #     | extend ClientIP=split(Client,":",0)
        #     | extend ClientPort=split(Client,":",1)
        # | project LDAPServer, TimeGenerated, ClientIP, ClientPort, StartingNode, Filter, SearchScope, AttributeSelection, ServerControls, VisitedEntries, ReturnedEntries, UsedIndexes, PagesReferenced, PagesReadFromDisk, PagesPreReadFromDisk, CleanPagesModified, DirtyPagesModified, SearchTimeMS, AttributesPreventingOptimization, User
  #    2. Output to CSV
  #    3. Put CSV in same directory as this script.
  #    4. Script will perform string replacement in TimeGenerated, ClientIP, ClientPort fields, then calls Excel to import resulting CSV, create pivot tables for common ldap search analysis scenarios. 
  # Note: Script requires 64bits Excel.
  #
  # DgrepLdapReader.ps1 
    #		Steps: 
    #   	1. Put downloaded Dgrep Log.*.csv to same directory as DgrepLdapReader.ps1
    #   	2. Run script

  # Script info:    https://docs.microsoft.com/en-us/troubleshoot/windows-server/identity/event1644reader-analyze-ldap-query-performance
    #   Latest:       https://github.com/mingchen-script/DgrepLdapReader.ps1
    # AD Schema:      https://docs.microsoft.com/en-us/windows/win32/adschema/active-directory-schema
    # AD Attributes:  https://docs.microsoft.com/en-us/windows/win32/adschema/attributes
#------Script variables block, modify to fit your needs ---------------------------------------------------------------------
  $g_ColorBar   = $True                 # Can set to $false to speed up excel import & reduce memory requirement. 
  $g_ColorScale = $True                 # Can set to $false to speed up excel import & reduce memory requirement. Color Scale requires '$g_ColorBar = $True' for color index. 
  $ErrorActionPreference = "SilentlyContinue"
function Set-PivotField { param ( $PivotField = $null, $Orientation = $null, $NumberFormat = $null, $Function = $null, $Calculation = $null, $Name = $null, $Group = $null )
    if ($null -ne $Orientation) {$PivotField.Orientation = $Orientation}
    if ($null -ne $NumberFormat) {$PivotField.NumberFormat = $NumberFormat}
    if ($null -ne $Function) {$PivotField.Function = $Function}
    if ($null -ne $Calculation) {$PivotField.Calculation = $Calculation}
    if ($null -ne $Name) {$PivotField.Name = $Name}
    if ($null -ne $Group) {($PivotField.DataRange.Item($group)).group($true,$true,1,($false, $true, $true, $true, $false, $false, $false)) | Out-Null}
}
function Set-PivotPageRows { param ( $Sheet = $null, $PivotTable = $null, $Page = $null, $Rows = $null  )
    $xlRowField   = 1 #XlPivotFieldOrientation 
    $xlPageField  = 3 #XlPivotFieldOrientation 
    Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$Page") -Orientation $xlPageField
    $i=0
    ($Rows).foreach({
      $i++
      If ($i -lt ($Rows).count) {Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$_") -Orientation $xlRowField}
      else {Set-PivotField -PivotField $Sheet.PivotTables("$PivotTable").PivotFields("$_") -Orientation $xlRowField -Group $i}
    })
}
function Set-TableFormats { param ( $Sheet = $null, $Table = $null, $ColumnWidth = $null, $label = $null, $Name = $null, $ColorScale = $null, $ColorBar = $null, $SortColumn = $null, $Hide = $null, $ColumnHiLite = $null, $NoteColumn = $null, $Note = $null )
  $Sheet.PivotTables("$Table").HasAutoFormat = $False
    $Column = 1
    $ColumnWidth.foreach({ $Sheet.columns.item($Column).columnwidth = $_
      $Column++
    })
    $Sheet.Application.ActiveWindow.SplitRow = 3
    $Sheet.Application.ActiveWindow.SplitColumn = 2
    $Sheet.Application.ActiveWindow.FreezePanes = $true
    $Sheet.Cells.Item(3,1) = $label
    $Sheet.Name = $Name
    if ($null -ne $SortColumn) {$null = $Sheet.Cells.Item($SortColumn,4).Sort($Sheet.Cells.Item($SortColumn,4),2)}
    if ($null -ne $Hide) {$Hide.foreach({($Sheet.PivotTables("$Table").PivotFields($_)).ShowDetail = $false})}
    if ($null -ne $ColumnHiLite) {
      $Sheet.Range("A4:"+[char]($sheet.UsedRange.Cells.Columns.count+64)+[string](($Sheet.UsedRange.Cells).Rows.count-1)).interior.Color = 16056319
      $ColumnHiLite.ForEach({$sheet.Range(($_+"3")).interior.ColorIndex = 37})
    }
    if (($null -ne $ColorBar) -and ($g_ColorBar -eq $true)) {
      $ColorRange='$'+$ColorBar+'$4:$'+$ColorBar+'$'+(($Sheet.UsedRange.Cells).Rows.Count-1)
      $null = $Sheet.Range($ColorRange).FormatConditions.AddDatabar()
    }
    if (($null -ne $ColorScale) -and ($g_ColorScale -eq $true)) {
      $ColorRange='$'+$ColorScale+'$4:$'+$ColorScale+'$'+(($Sheet.UsedRange.Cells).Rows.Count-1)
      $null = $Sheet.Range($ColorRange).FormatConditions.AddColorScale(3)
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(1).type = 1
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(1).FormatColor.Color = 8109667
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(2).FormatColor.Color = 8711167
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(3).type = 2 
        $Sheet.Range($ColorRange).FormatConditions.item(2).ColorScaleCriteria.item(3).FormatColor.Color = 7039480
    }
    $Sheet.Cells.Item(1,$NoteColumn)= "[More Info]" #--Add log info
      $null = $Sheet.Cells.Item(1,$NoteColumn).addcomment()
      $null = $Sheet.Cells.Item(1,$NoteColumn).comment.text($Note)
      $Sheet.Cells.Item(1,$NoteColumn).comment.shape.textframe.Autosize = $true
  }
#------Main---------------------------------
$ScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
  $TotalSteps = ((Get-ChildItem -Path $ScriptPath -Filter 'Log*.csv').count)+10
  $Step=1
  $TimeStamp = "{0:yyyy-MM-dd_hh-mm-ss_tt}" -f (Get-Date)
  #---------Find logs's time range Info----------
  Write-Progress -Activity "Checking CSV timeRange" -PercentComplete (($Step++/$TotalSteps)*100)
  $OldestTimeStamp = $NewestTimeStamp = $LogsInfo = $null
  (Get-ChildItem -Path $ScriptPath\* -include ('Logs*.csv') ).foreach({
    $FirstTimeStamp = [DateTime]((Get-Content $_ -Tail 1) -split ',' | Select-Object -skip 1 -first 1 | ForEach-Object { $_ -replace '"',$null})
    $LastTimeStamp = [DateTime]((Get-Content $_ -Head 2) -split ',' | Select-Object -skip 21 -first 1 | ForEach-Object { $_ -replace '"',$null})
      if ($OldestTimeStamp -eq $null) { $OldestTimeStamp = $NewestTimeStamp = $FirstTimeStamp }
      If ($OldestTimeStamp -gt $FirstTimeStamp) {$OldestTimeStamp = $FirstTimeStamp }
      If ($NewestTimeStamp -lt $LastTimeStamp) {$NewestTimeStamp = $LastTimeStamp }
      $LogsInfo = $LogsInfo + ($_.name+"`n   "+$FirstTimeStamp+' ~ '+$LastTimeStamp+"`t   Log range = "+($LastTimeStamp-$FirstTimeStamp).Days+" Days "+($LastTimeStamp-$FirstTimeStamp).Hours+" Hours "+($LastTimeStamp-$FirstTimeStamp).Minutes+" min "+($LastTimeStamp-$FirstTimeStamp).Seconds+" sec. ("+((Get-Content $_ | Measure-Object -line).lines-1)+" Events.)`n`n")
  })
    $LogTimeRange = ($NewestTimeStamp-$OldestTimeStamp)
    $LogRangeText += ("Script info:`n   https://docs.microsoft.com/en-us/troubleshoot/windows-server/identity/event1644reader-analyze-ldap-query-performance`n") 
    $LogRangeText += ("Github latest download:`n   https://github.com/mingchen-script/DgrepLdapReader`n`n") 
    $LogRangeText += ("AD Schema:`n   https://docs.microsoft.com/en-us/windows/win32/adschema/active-directory-schema`n") 
    $LogRangeText += ("AD Attributes:`n   https://docs.microsoft.com/en-us/windows/win32/adschema/attributes`n`n") 
    $LogRangeText += ("#-------------------------------`n  [Overall EventRange]: "+$OldestTimeStamp+' ~ '+$NewestTimeStamp+"`n  [Overall TimeRange]: "+$LogTimeRange.Days+' Days '+$LogTimeRange.Hours+' Hours '+$LogTimeRange.Minutes+' Minutes '+$LogTimeRange.Seconds+" Seconds `n`n") + ($LogsInfo -replace "$TimeStamp-Temp1644-")
#-----Combine CSV(s) into one for faster Excel import
  $OutTitle1 = 'DGrepLdapSearches'
  $OutFile1 = "$ScriptPath\$TimeStamp-$OutTitle1.csv"
  Write-Progress -Activity "Generating $OutTitle1" -PercentComplete (($Step++/$TotalSteps)*100)
    (Get-ChildItem -Path $ScriptPath -Filter "Logs*.csv" | Select-Object -ExpandProperty FullName).foreach({
      $1644s=Import-Csv $_
        foreach ($1644 in $1644s) {
          $1644.TimeGenerated=[DateTime]($1644.TimeGenerated)
          $1644.ClientIP=$1644.ClientIP.replace('["','')
          $1644.ClientIP=$1644.ClientIP.replace('"]','')
          $1644.ClientPort=$1644.ClientPort.replace('["','')
          $1644.ClientPort=$1644.ClientPort.replace('"]','')
        }
    })
    $1644s | Export-Csv $OutFile1 -NoTypeInformation -Append
    #$null = Get-ChildItem -Path $ScriptPath -Filter "Logs*.csv" | Remove-Item
#----Excel COM variables-------------------------------------------------------------------
  $fmtNumber  = "###,###,###,###,###"
  $fmtPercent = "#0.00%"
  $xlDataField  = 4 #XlPivotFieldOrientation 
  $xlAverage    = -4106 #XlConsolidationFunction
  $xlSum        = -4157 #XlConsolidationFunction 
  $xlPercentOfTotal = 8 #XlPivotFieldCalculation 
#-------Import to Excel
If (Test-Path $OutFile1) { 
  $Excel = New-Object -ComObject excel.application
  Write-Progress -Activity "Import to Excel $OutTitle1" -PercentComplete (($Step++/$TotalSteps)*100)
    # $Excel.visible = $true
    $Excel.Workbooks.OpenText("$OutFile1")
    $Sheet0 = $Excel.Workbooks[1].Worksheets[1]
      $Sheet0.Application.ActiveWindow.SplitRow=1  
      $Sheet0.Application.ActiveWindow.FreezePanes = $true
      $null = $Sheet0.Columns.AutoFit() = $Sheet0.Range("A1").AutoFilter()
        ("C","D","J","K","L").ForEach({$Sheet0.Columns.Item($_).columnwidth = 70})
        ("E","F","H","M","N","O","P","Q","R").ForEach({$Sheet0.Columns.Item($_).numberformat = $fmtNumber})
        $Sheet0.Columns.Item("B").numberformat = "m/d/yyyy h:mm:s AM/PM"
      $Sheet0.Name = $OutTitle1
      $null = $Sheet0.ListObjects.Add(1, $Sheet0.Application.ActiveCell.CurrentRegion, $null ,0)
    #----Pivot Table 1-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount StartingNode Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet1 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet1!R1C1")
      Set-PivotPageRows -Sheet $sheet1 -PivotTable "PivotTable1" -Page "LDAPServer" -Rows ("StartingNode","Filter","ClientIP","TimeGenerated")
        Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" 
        Set-PivotField -PivotField $Sheet1.PivotTables("PivotTable1").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet1 -Table "PivotTable1" -ColumnWidth (60,12,14,12,14) -label 'StartingNode grouping' -Name '1.TopCount StartingNode' -SortColumn 4 -Hide ('ClientIP','Filter','StartingNode') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
      #----Pivot Table 2-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount IP Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet2 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet2!R1C1")
      Set-PivotPageRows -Sheet $sheet2 -PivotTable "PivotTable2" -Page "LDAPServer" -Rows ("ClientIP","Filter","TimeGenerated")
        Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" 
        Set-PivotField -PivotField $Sheet2.PivotTables("PivotTable2").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet2 -Table "PivotTable2" -ColumnWidth (60,12,19,12) -label 'IP grouping' -Name '2.TopCount IP' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 3-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopCount Filters Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet3 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet3!R1C1")
      Set-PivotPageRows -Sheet $sheet3 -PivotTable "PivotTable3" -Page "LDAPServer" -Rows ("Filter","ClientIP","TimeGenerated")
        Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlAverage -Name "AvgSearchTime" 
        Set-PivotField -PivotField $Sheet3.PivotTables("PivotTable3").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet3 -Table "PivotTable3" -ColumnWidth (70,12,19,12) -label 'Filter grouping' -Name '3.TopCount Filters' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 4-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopTime IP Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet4 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet4!R1C1")
      Set-PivotPageRows -Sheet $sheet4 -PivotTable "PivotTable4" -Page "LDAPServer" -Rows ("ClientIP","Filter","TimeGenerated")
        Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" 
        Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet4.PivotTables("PivotTable4").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet4 -Table "PivotTable4" -ColumnWidth (50,21,12,19) -label 'IP grouping' -Name '4.TopTime IP' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 5-------------------------------------------------------------------
    Write-Progress -Activity "Creating TopTime Filter Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet5 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet5!R1C1")
      Set-PivotPageRows -Sheet $sheet5 -PivotTable "PivotTable5" -Page "LDAPServer" -Rows ("Filter","ClientIP","TimeGenerated")
        Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" 
        Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet5.PivotTables("PivotTable5").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet5 -Table "PivotTable5" -ColumnWidth (70,21,12,19) -label 'Filter grouping' -Name '5.TopTime Filter' -SortColumn 4 -Hide ('ClientIP','Filter') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 6-------------------------------------------------------------------
    Write-Progress -Activity "Creating Top Users Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet6 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet6!R1C1")
      Set-PivotPageRows -Sheet $Sheet6 -PivotTable "PivotTable6" -Page "LDAPServer" -Rows ("User","ClientIP","Filter")
        Set-PivotField -PivotField $Sheet6.PivotTables("PivotTable6").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" 
        Set-PivotField -PivotField $Sheet6.PivotTables("PivotTable6").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet6.PivotTables("PivotTable6").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet6 -Table "PivotTable6" -ColumnWidth (70,21,12,19) -label 'Filter grouping' -Name '6.TopTime User' -SortColumn 4 -Hide ('Filter','ClientIP','User') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #----Pivot Table 7-------------------------------------------------------------------
    Write-Progress -Activity "Creating Top Filter Pivot table" -PercentComplete (($Step++/$TotalSteps)*100)
    $Sheet7 = $Excel.Workbooks[1].Worksheets.add()
    $null = ($Excel.Workbooks[1].PivotCaches().Create(1,"$OutTitle1!R1C1:R$($Sheet0.UsedRange.Rows.count)C$($Sheet0.UsedRange.Columns.count)",5)).CreatePivotTable("Sheet7!R1C1")
      Set-PivotPageRows -Sheet $Sheet7 -PivotTable "PivotTable7" -Page "LDAPServer" -Rows ("AttributesPreventingOptimization","Filter","ClientIP")
        Set-PivotField -PivotField $Sheet7.PivotTables("PivotTable7").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtNumber -Function $xlSum -Name "Total SearchTime" 
        Set-PivotField -PivotField $Sheet7.PivotTables("PivotTable7").PivotFields("ClientIP") -Orientation $xlDataField -NumberFormat  $fmtNumber -Name "Search Count" 
        Set-PivotField -PivotField $Sheet7.PivotTables("PivotTable7").PivotFields("SearchTimeMS") -Orientation $xlDataField -NumberFormat  $fmtPercent -Calculation $xlPercentOfTotal -Name "%GrandTotal"
      Set-TableFormats -Sheet $Sheet7 -Table "PivotTable7" -ColumnWidth (70,21,12,19) -label 'Attributes Preventing Optimization (Need Index)' -Name '7.Attributes Need Optimization' -SortColumn 4 -Hide ('ClientIP','Filter','AttributesPreventingOptimization') -ColumnHiLite ('B','D') -ColorBar 'D' -ColorScale 'D' -NoteColumn 'C' -Note $LogRangeText
    #---General Tab Operations-------------------------------------------------------------------
    ($Sheet1,$Sheet2,$Sheet3).ForEach{$_.Tab.ColorIndex = 35}
    ($Sheet4,$Sheet5).ForEach{$_.Tab.ColorIndex = 36}
    ($Sheet6,$Sheet7).ForEach{$_.Tab.ColorIndex = 37}
      $WorkSheetNames = New-Object System.Collections.ArrayList  #---Sort by sheetName-
      foreach($WorkSheet in $Excel.Workbooks[1].Worksheets) { $null = $WorkSheetNames.add($WorkSheet.Name) }
        $null = $WorkSheetNames.Sort()
        For ($i=0; $i -lt $WorkSheetNames.Count-1; $i++){ ($Excel.Workbooks[1].Worksheets.Item($WorkSheetNames[$i])).Move($Excel.Workbooks[1].Worksheets.Item($i+1)) }
    $Sheet1.Activate()
    $Excel.Workbooks[1].SaveAs($ScriptPath+'\'+$TimeStamp+'-'+$OutTitle1,51)
    Remove-Item "$ScriptPath\$TimeStamp-$OutTitle1.csv"
    $Excel.visible = $true
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
      # Stop-process -Name Excel 
} else {
	Write-Host 'No DGrep CSV found. Please confirm Log*.csv is in the script directory.' -ForegroundColor Red
}

# Excel数据分析脚本
# 作者：元宝
# 日期：2026-03-25
# 版本：1.8

# 获取脚本文件本身的目录路径
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "脚本所在目录: $scriptDir" -ForegroundColor Cyan

# 导入必要的模块
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 在脚本所在目录中查找Excel文件
$excelFiles = Get-ChildItem -Path $scriptDir -Filter "OTDR*.xlsx"
$outputData = @()
$allPNData = @{}

# 创建输出目录
$outputDir = Join-Path $scriptDir "AnalysisResults_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

Write-Host "开始处理Excel文件..." -ForegroundColor Green
Write-Host "找到 $($excelFiles.Count) 个Excel文件" -ForegroundColor Yellow

# 如果没有找到文件，显示提示
if ($excelFiles.Count -eq 0) {
    Write-Host "未找到Excel文件！" -ForegroundColor Red
    Write-Host "请确保脚本同目录下有OTDR开头的Excel文件" -ForegroundColor Yellow
    Write-Host "`n按任意键退出..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 定义要处理的行位置
$iRows = @(3,12,21,30,39,48,57,66,75)  # 每组起始行

# 函数：计算工作表指定位置的平均值
function Get-SheetColumnAverages {
    param(
        [object]$Sheet,
        [int]$Column,
        [int]$PositionCount = 8
    )
    
    $averages = @()
    for ($pos = 0; $pos -lt $PositionCount; $pos++) {
        $values = @()
        foreach ($startRow in $iRows) {
            $row = $startRow + $pos
            $cellValue = $Sheet.Cells.Item($row, $Column).Value2
            if ($cellValue -ne $null) {
                $values += [double]$cellValue
            }
        }
        $averages += if ($values.Count -gt 0) { 
            [Math]::Round(($values | Measure-Object -Average).Average, 6) 
        } else { 0 }
    }
    return $averages
}

# 函数：处理PN 997的Reflectance数据
function Process-PN997Reflectance {
    param(
        [object]$Workbook
    )
    
    $powerSheets = @{
        "LOW" = "Reflectance data (case7)_LOW"
        "MID" = "Reflectance data (case7)_MID"
        "HIGH" = "Reflectance data (case7)_HIGH"
    }
    
    $powerData = @{}
    $allPositions = @{}  # 存储所有功率的合并数据
    
    # 初始化位置数组
    for ($i = 0; $i -lt 8; $i++) {
        $allPositions[$i] = @{ G = @(); I = @(); J = @() }
    }
    
    $processedSheets = 0
    
    foreach ($power in $powerSheets.Keys) {
        $sheetName = $powerSheets[$power]
        $sheet = $Workbook.Worksheets | Where-Object { $_.Name -eq $sheetName }
        
        if (-not $sheet) { continue }
        
        Write-Host "    - 处理$sheetName..." -ForegroundColor DarkGray
        
        $powerData[$power] = @{
            G = @()  # 距离
            I = @()  # 插损
            J = @()  # 反射
        }
        
        # 处理每个位置
        for ($pos = 0; $pos -lt 8; $pos++) {
            $gValues = @()
            $iValues = @()
            $jValues = @()
            
            # 收集当前功率当前位置的所有行数据
            foreach ($startRow in $iRows) {
                $row = $startRow + $pos
                $gVal = $sheet.Cells.Item($row, 7).Value2
                $iVal = $sheet.Cells.Item($row, 9).Value2
                $jVal = $sheet.Cells.Item($row, 10).Value2
                
                if ($gVal -ne $null) { 
                    $gValues += [double]$gVal
                    $allPositions[$pos].G += [double]$gVal
                }
                if ($iVal -ne $null) { 
                    $iValues += [double]$iVal
                    $allPositions[$pos].I += [double]$iVal
                }
                if ($jVal -ne $null) { 
                    $jValues += [double]$jVal
                    $allPositions[$pos].J += [double]$jVal
                }
            }
            
            # 计算当前功率当前位置的平均值
            $powerData[$power].G += if ($gValues.Count -gt 0) { 
                [Math]::Round(($gValues | Measure-Object -Average).Average, 6) 
            } else { 0 }
            
            $powerData[$power].I += if ($iValues.Count -gt 0) { 
                [Math]::Round(($iValues | Measure-Object -Average).Average, 6) 
            } else { 0 }
            
            $powerData[$power].J += if ($jValues.Count -gt 0) { 
                [Math]::Round(($jValues | Measure-Object -Average).Average, 6) 
            } else { 0 }
        }
        
        $processedSheets++
        Write-Host "    √ $sheetName处理完成" -ForegroundColor DarkGreen
    }
    
    if ($processedSheets -eq 0) { return $null }
    
    # 计算所有功率的总平均值
    $finalG = @()
    $finalI = @()
    $finalJ = @()
    
    for ($i = 0; $i -lt 8; $i++) {
        $finalG += if ($allPositions[$i].G.Count -gt 0) {
            [Math]::Round(($allPositions[$i].G | Measure-Object -Average).Average, 6)
        } else { 0 }
        
        $finalI += if ($allPositions[$i].I.Count -gt 0) {
            [Math]::Round(($allPositions[$i].I | Measure-Object -Average).Average, 6)
        } else { 0 }
        
        $finalJ += if ($allPositions[$i].J.Count -gt 0) {
            [Math]::Round(($allPositions[$i].J | Measure-Object -Average).Average, 6)
        } else { 0 }
    }
    
    return @{
        PowerData = $powerData
        TotalAverages = @{
            G = $finalG
            I = $finalI
            J = $finalJ
        }
        SheetCount = $processedSheets
    }
}

# 处理每个Excel文件
foreach ($file in $excelFiles) {
    Write-Host "处理文件: $($file.Name)" -ForegroundColor Cyan
    
    # 从文件名中提取PN号
    $fileName = $file.BaseName
    $pattern = "OTDR[_\s]*(\d+)_[A-Za-z]\d+_\d{4}-\d{2}-\d{2}_TestReport"
    
    if ($fileName -match $pattern) {
        $pn = $matches[1]
        $filePath = $file.FullName
        
        Write-Host "  提取到PN号: $pn" -ForegroundColor Green
        
        # 创建Excel COM对象
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        try {
            $workbook = $excel.Workbooks.Open($filePath)
            
            # 初始化PN数据对象
            $pnData = [PSCustomObject]@{
                PN = $pn
                FileName = $file.Name
                LossData_I = @()  # case8 I列 (插损)
                LossData_G = @()  # case8 G列 (距离)
                ReflectanceData_G = @()  # case7 G列 (距离)
                ReflectanceData_I = @()  # case7 I列 (插损)
                ReflectanceData_J = @()  # case7 J列 (反射)
                PM_Data = [PSCustomObject]@{
                    Average = $null
                    Min = $null
                    Max = $null
                }
            }
            
            # 处理Loss data(case8) sheet
            Write-Host "  - 处理Loss data(case8)..." -ForegroundColor Gray
            $lossSheet = $workbook.Worksheets | Where-Object {$_.Name -eq "Loss data(case8)"}
            if ($lossSheet) {
                $pnData.LossData_I = Get-SheetColumnAverages -Sheet $lossSheet -Column 9
                $pnData.LossData_G = Get-SheetColumnAverages -Sheet $lossSheet -Column 7
                Write-Host "  √ Loss data(case8)处理完成" -ForegroundColor Green
            } else {
                Write-Host "  × 未找到Loss data(case8)工作表" -ForegroundColor Red
            }
            
            # 处理Reflectance data
            Write-Host "  - 处理Reflectance data..." -ForegroundColor Gray
            
            if ($pn -eq "1831781997") {
                $pn997Data = Process-PN997Reflectance -Workbook $workbook
                if ($pn997Data) {
                    $pnData.ReflectanceData_G = $pn997Data.TotalAverages.G
                    $pnData.ReflectanceData_I = $pn997Data.TotalAverages.I
                    $pnData.ReflectanceData_J = $pn997Data.TotalAverages.J
                    $pnData | Add-Member -NotePropertyName "ReflectancePowerData" -NotePropertyValue $pn997Data.PowerData
                    Write-Host "  √ Reflectance data处理完成 (处理了$($pn997Data.SheetCount)个功率级别)" -ForegroundColor Green
                } else {
                    Write-Host "  × 未找到任何Reflectance data工作表" -ForegroundColor Red
                }
            } else {
                # 处理普通Reflectance data (case7)
                $refSheet = $workbook.Worksheets | Where-Object {$_.Name -eq "Reflectance data (case7)"}
                if (-not $refSheet) {
                    $refSheet = $workbook.Worksheets | Where-Object {$_.Name -match "Reflectance.*case7"}
                }
                
                if ($refSheet) {
                    $pnData.ReflectanceData_G = Get-SheetColumnAverages -Sheet $refSheet -Column 7
                    $pnData.ReflectanceData_I = Get-SheetColumnAverages -Sheet $refSheet -Column 9
                    $pnData.ReflectanceData_J = Get-SheetColumnAverages -Sheet $refSheet -Column 10
                    Write-Host "  √ Reflectance data (case7)处理完成" -ForegroundColor Green
                } else {
                    Write-Host "  × 未找到Reflectance data (case7)工作表" -ForegroundColor Yellow
                }
            }
            
            # 处理Reflectance data(PM)
            Write-Host "  - 处理Reflectance data(PM)..." -ForegroundColor Gray
            $pmSheet = $workbook.Worksheets | Where-Object {$_.Name -eq "Reflectance data(PM)"}
            if ($pmSheet) {
                $values = @()
                for ($row = 2; $row -le 28; $row++) {
                    $cellValue = $pmSheet.Cells.Item($row, 24).Value2
                    if ($cellValue -ne $null) {
                        $values += [double]$cellValue
                    }
                }
                
                if ($values.Count -gt 0) {
                    $stats = $values | Measure-Object -Average -Minimum -Maximum
                    $pnData.PM_Data.Average = [Math]::Round($stats.Average, 6)
                    $pnData.PM_Data.Min = [Math]::Round($stats.Minimum, 6)
                    $pnData.PM_Data.Max = [Math]::Round($stats.Maximum, 6)
                    Write-Host "  √ PM Data处理完成: 平均值=$([Math]::Round($stats.Average,6))" -ForegroundColor Green
                } else {
                    Write-Host "  × 未在Reflectance data(PM)中找到数据" -ForegroundColor Yellow
                }
            } else {
                Write-Host "  × 未找到Reflectance data(PM)工作表" -ForegroundColor Yellow
            }
            
            # 将PN数据添加到集合
            $allPNData[$pn] = $pnData
            $outputData += $pnData
            Write-Host "  √ 文件处理完成: $pn" -ForegroundColor Green
            
        } catch {
            Write-Host "  × 处理文件出错: $($_.Exception.Message)" -ForegroundColor Red
        } finally {
            # 清理Excel对象
            try {
                if ($workbook) { 
                    $workbook.Close($false) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                }
                if ($excel) { 
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                }
            } catch { }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    } else {
        Write-Host "× 跳过文件（格式不匹配）: $fileName" -ForegroundColor Yellow
    }
}

# 如果没有成功处理任何文件，退出
if ($allPNData.Count -eq 0) {
    Write-Host "`n没有成功处理任何文件！" -ForegroundColor Red
    Write-Host "按任意键退出..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 计算差异分析
Write-Host "`n开始计算数据差异..." -ForegroundColor Green
$pns = $allPNData.Keys | Sort-Object
$pnCount = $pns.Count

# 函数：比较数据集差异
function Compare-DataSets {
    param(
        [hashtable]$AllData,
        [string]$PropertyPath,
        [string]$CategoryName
    )
    
    $results = @()
    for ($i = 0; $i -lt 8; $i++) {
        $distanceData = @()
        foreach ($pn in $pns) {
            $data = $AllData[$pn]
            $value = $data
            foreach ($prop in $PropertyPath.Split('.')) {
                $value = $value.$prop
            }
            
            if ($value -and $i -lt $value.Count) {
                $distanceData += [PSCustomObject]@{
                    PN = $pn
                    Value = $value[$i]
                }
            }
        }
        
        if ($distanceData.Count -gt 1) {
            $valuesArray = $distanceData.Value
            $averageValue = [Math]::Round(($valuesArray | Measure-Object -Average).Average, 6)
            
            $distances = $distanceData | ForEach-Object {
                [PSCustomObject]@{
                    PN = $_.PN
                    Distance = [Math]::Round([Math]::Abs($_.Value - $averageValue), 6)
                }
            }
            
            $results += [PSCustomObject]@{
                Distance = "距离_$($i+1)"
                Index = $i+1
                ClosestPoint = $averageValue
                PNs = ($distanceData.PN -join ", ")
                Values = ($distanceData.Value | ForEach-Object { [Math]::Round($_, 6) }) -join ", "
                Distances = ($distances.Distance -join ", ")
                DistanceDetails = $distances
            }
        }
    }
    
    if ($results.Count -gt 0) {
        return [PSCustomObject]@{
            Category = $CategoryName
            Results = $results
        }
    }
    return $null
}

# 执行差异分析
$analysisResults = @()
if ($pnCount -ge 2) {
    $categories = @(
        @{ Path = "LossData_I"; Name = "Loss Data (case8) I列(插损)" },
        @{ Path = "LossData_G"; Name = "Loss Data (case8) G列(距离)" },
        @{ Path = "ReflectanceData_G"; Name = "Reflectance Data (case7) G列(距离)" },
        @{ Path = "ReflectanceData_I"; Name = "Reflectance Data (case7) I列(插损)" },
        @{ Path = "ReflectanceData_J"; Name = "Reflectance Data (case7) J列(反射)" }
    )
    
    foreach ($category in $categories) {
        Write-Host "分析$($category.Name)差异..." -ForegroundColor Yellow
        $result = Compare-DataSets -AllData $allPNData -PropertyPath $category.Path -CategoryName $category.Name
        if ($result) {
            $analysisResults += $result
            Write-Host "  √ 分析完成" -ForegroundColor Green
        }
    }
} else {
    Write-Host "只有1个PN号，无法进行差异分析" -ForegroundColor Yellow
}

# 比较PM数据
Write-Host "`n分析PM Data差异..." -ForegroundColor Yellow
$pmResults = $pns | ForEach-Object {
    $data = $allPNData[$_]
    if ($data.PM_Data.Average -ne $null) {
        [PSCustomObject]@{
            PN = $data.PN
            Average = $data.PM_Data.Average
            Min = $data.PM_Data.Min
            Max = $data.PM_Data.Max
        }
    }
}

# 创建输出Excel文件
Write-Host "`n生成输出文件..." -ForegroundColor Green

try {
    $outputExcel = New-Object -ComObject Excel.Application
    $outputExcel.Visible = $false
    $outputExcel.DisplayAlerts = $false
    $outputWorkbook = $outputExcel.Workbooks.Add()
    
    # 1. 创建Summary工作表
    $summarySheet = $outputWorkbook.Worksheets.Item(1)
    $summarySheet.Name = "数据汇总"
    
    $headers = @("PN号", "文件名", "Loss I列(插损)", "Loss G列(距离)", "Reflectance G(距离)", "Reflectance I(插损)", "Reflectance J(反射)", "PM平均值", "PM最小值", "PM最大值")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $summarySheet.Cells.Item(1, $i+1) = $headers[$i]
        $summarySheet.Cells.Item(1, $i+1).Font.Bold = $true
        $summarySheet.Cells.Item(1, $i+1).Interior.ColorIndex = 15
    }
    
    $row = 2
    foreach ($pn in $pns) {
        $data = $allPNData[$pn]
        $summarySheet.Cells.Item($row, 1) = $data.PN
        $summarySheet.Cells.Item($row, 2) = $data.FileName
        
        $columns = @(
            $data.LossData_I,
            $data.LossData_G,
            $data.ReflectanceData_G,
            $data.ReflectanceData_I,
            $data.ReflectanceData_J
        )
        
        for ($col = 0; $col -lt 5; $col++) {
            if ($columns[$col].Count -gt 0) {
                $summarySheet.Cells.Item($row, $col+3) = ($columns[$col] | ForEach-Object { [Math]::Round($_, 6) }) -join ", "
            }
        }
        
        if ($data.PM_Data.Average -ne $null) {
            $summarySheet.Cells.Item($row, 8) = [Math]::Round($data.PM_Data.Average, 6)
            $summarySheet.Cells.Item($row, 9) = [Math]::Round($data.PM_Data.Min, 6)
            $summarySheet.Cells.Item($row, 10) = [Math]::Round($data.PM_Data.Max, 6)
        }
        
        $row++
    }
    $summarySheet.UsedRange.EntireColumn.AutoFit() | Out-Null
    
    # 2. 创建PN详细数据工作表
    $detailSheet = $outputWorkbook.Worksheets.Add()
    $detailSheet.Name = "PN详细数据"
    
    $row = 1
    foreach ($pn in $pns) {
        $data = $allPNData[$pn]
        
        $detailSheet.Cells.Item($row, 1) = "PN: " + $data.PN
        $detailSheet.Cells.Item($row, 1).Font.Bold = $true
        $detailSheet.Cells.Item($row, 1).Font.Size = 12
        $detailSheet.Cells.Item($row, 1).Interior.ColorIndex = 15
        $row++
        
        $detailSheet.Cells.Item($row, 1) = "文件: " + $data.FileName
        $row += 2
        
        # 输出每个数据类别
        $dataCategories = @(
            @{ Name = "Loss Data (case8) - I列(插损)"; Data = $data.LossData_I },
            @{ Name = "Loss Data (case8) - G列(距离)"; Data = $data.LossData_G },
            @{ Name = "Reflectance Data (case7) - G列(距离)"; Data = $data.ReflectanceData_G },
            @{ Name = "Reflectance Data (case7) - I列(插损)"; Data = $data.ReflectanceData_I },
            @{ Name = "Reflectance Data (case7) - J列(反射)"; Data = $data.ReflectanceData_J }
        )
        
        foreach ($category in $dataCategories) {
            if ($category.Data.Count -gt 0) {
                $detailSheet.Cells.Item($row, 1) = $category.Name
                $detailSheet.Cells.Item($row, 1).Font.Bold = $true
                $row++
                
                $detailSheet.Cells.Item($row, 1) = "距离"
                $detailSheet.Cells.Item($row, 2) = "平均值"
                $row++
                
                for ($i = 1; $i -le 8; $i++) {
                    $detailSheet.Cells.Item($row, 1) = "距离 $i"
                    if (($i-1) -lt $category.Data.Count) {
                        $detailSheet.Cells.Item($row, 2) = [Math]::Round($category.Data[$i-1], 6)
                    }
                    $row++
                }
                $row += 2
            }
        }
        
        # 特殊处理PN 997的功率数据
        if ($pn -eq "1831781997" -and $data.ReflectancePowerData) {
            $detailSheet.Cells.Item($row, 1) = "PN 997 功率分析"
            $detailSheet.Cells.Item($row, 1).Font.Bold = $true
            $detailSheet.Cells.Item($row, 1).Font.Color = 10498160
            $row += 2
            
            $powers = @("LOW", "MID", "HIGH")
            $colors = @(255, 32768, 12611584)  # 红、绿、蓝
            
            for ($p = 0; $p -lt $powers.Count; $p++) {
                $power = $powers[$p]
                if ($data.ReflectancePowerData[$power]) {
                    $detailSheet.Cells.Item($row, 1) = "功率级别: $power"
                    $detailSheet.Cells.Item($row, 1).Font.Bold = $true
                    $detailSheet.Cells.Item($row, 1).Font.Color = $colors[$p]
                    $row++
                    
                    $detailSheet.Cells.Item($row, 1) = "距离"
                    $detailSheet.Cells.Item($row, 2) = "G列(距离)"
                    $detailSheet.Cells.Item($row, 3) = "I列(插损)"
                    $detailSheet.Cells.Item($row, 4) = "J列(反射)"
                    $row++
                    
                    for ($i = 1; $i -le 8; $i++) {
                        $detailSheet.Cells.Item($row, 1) = "距离 $i"
                        if (($i-1) -lt $data.ReflectancePowerData[$power].G.Count) {
                            $detailSheet.Cells.Item($row, 2) = [Math]::Round($data.ReflectancePowerData[$power].G[$i-1], 6)
                        }
                        if (($i-1) -lt $data.ReflectancePowerData[$power].I.Count) {
                            $detailSheet.Cells.Item($row, 3) = [Math]::Round($data.ReflectancePowerData[$power].I[$i-1], 6)
                        }
                        if (($i-1) -lt $data.ReflectancePowerData[$power].J.Count) {
                            $detailSheet.Cells.Item($row, 4) = [Math]::Round($data.ReflectancePowerData[$power].J[$i-1], 6)
                        }
                        $row++
                    }
                    $row += 2
                }
            }
            
            $detailSheet.Cells.Item($row, 1) = "所有功率总平均值"
            $detailSheet.Cells.Item($row, 1).Font.Bold = $true
            $detailSheet.Cells.Item($row, 1).Font.Color = 10498160
            $row++
            
            $detailSheet.Cells.Item($row, 1) = "距离"
            $detailSheet.Cells.Item($row, 2) = "G列(距离)"
            $detailSheet.Cells.Item($row, 3) = "I列(插损)"
            $detailSheet.Cells.Item($row, 4) = "J列(反射)"
            $row++
            
            for ($i = 1; $i -le 8; $i++) {
                $detailSheet.Cells.Item($row, 1) = "距离 $i"
                if (($i-1) -lt $data.ReflectanceData_G.Count) {
                    $detailSheet.Cells.Item($row, 2) = [Math]::Round($data.ReflectanceData_G[$i-1], 6)
                }
                if (($i-1) -lt $data.ReflectanceData_I.Count) {
                    $detailSheet.Cells.Item($row, 3) = [Math]::Round($data.ReflectanceData_I[$i-1], 6)
                }
                if (($i-1) -lt $data.ReflectanceData_J.Count) {
                    $detailSheet.Cells.Item($row, 4) = [Math]::Round($data.ReflectanceData_J[$i-1], 6)
                }
                $row++
            }
            $row += 3
        }
        
        # PM Data
        if ($data.PM_Data.Average -ne $null) {
            $detailSheet.Cells.Item($row, 1) = "Reflectance Data(PM)"
            $detailSheet.Cells.Item($row, 1).Font.Bold = $true
            $row++
            
            $detailSheet.Cells.Item($row, 1) = "平均值"
            $detailSheet.Cells.Item($row, 2) = [Math]::Round($data.PM_Data.Average, 6)
            $row++
            
            $detailSheet.Cells.Item($row, 1) = "最小值"
            $detailSheet.Cells.Item($row, 2) = [Math]::Round($data.PM_Data.Min, 6)
            $row++
            
            $detailSheet.Cells.Item($row, 1) = "最大值"
            $detailSheet.Cells.Item($row, 2) = [Math]::Round($data.PM_Data.Max, 6)
            $row += 3
        }
    }
    
    $detailSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
    
    # 3. 创建差异分析工作表
    if ($analysisResults.Count -gt 0) {
        $compareSheet = $outputWorkbook.Worksheets.Add()
        $compareSheet.Name = "差异分析"
        
        $compareHeaders = @("数据类别", "距离", "平均值", "PN列表", "各PN数值", "到平均值距离")
        for ($i = 0; $i -lt $compareHeaders.Count; $i++) {
            $compareSheet.Cells.Item(1, $i+1) = $compareHeaders[$i]
            $compareSheet.Cells.Item(1, $i+1).Font.Bold = $true
            $compareSheet.Cells.Item(1, $i+1).Interior.ColorIndex = 15
        }
        
        $row = 2
        foreach ($analysis in $analysisResults) {
            foreach ($result in $analysis.Results) {
                $compareSheet.Cells.Item($row, 1) = $analysis.Category
                $compareSheet.Cells.Item($row, 2) = $result.Distance
                $compareSheet.Cells.Item($row, 3) = $result.ClosestPoint
                $compareSheet.Cells.Item($row, 4) = $result.PNs
                $compareSheet.Cells.Item($row, 5) = $result.Values
                $compareSheet.Cells.Item($row, 6) = $result.Distances
                $row++
            }
            $row++  # 添加空行
        }
        $compareSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
    }
    
    # 4. 创建PM数据工作表
    if ($pmResults.Count -gt 0) {
        $pmSheet = $outputWorkbook.Worksheets.Add()
        $pmSheet.Name = "PM数据分析"
        
        $pmHeaders = @("PN号", "平均值", "最小值", "最大值")
        for ($i = 0; $i -lt $pmHeaders.Count; $i++) {
            $pmSheet.Cells.Item(1, $i+1) = $pmHeaders[$i]
            $pmSheet.Cells.Item(1, $i+1).Font.Bold = $true
            $pmSheet.Cells.Item(1, $i+1).Interior.ColorIndex = 15
        }
        
        $row = 2
        foreach ($pmResult in $pmResults) {
            $pmSheet.Cells.Item($row, 1) = $pmResult.PN
            $pmSheet.Cells.Item($row, 2) = [Math]::Round($pmResult.Average, 6)
            $pmSheet.Cells.Item($row, 3) = [Math]::Round($pmResult.Min, 6)
            $pmSheet.Cells.Item($row, 4) = [Math]::Round($pmResult.Max, 6)
            $row++
        }
        $pmSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
    }
    
    # 保存文件
    $outputPath = Join-Path $outputDir "OTDR数据分析报告_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $outputWorkbook.SaveAs($outputPath)
    
    Write-Host "`n√ 分析完成！" -ForegroundColor Green
    Write-Host "输出文件已保存到: $outputPath" -ForegroundColor Yellow
    
    # 释放COM对象
    $outputWorkbook.Close($true)
    $outputExcel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outputExcel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    
} catch {
    Write-Host "× 保存输出文件时出错: $($_.Exception.Message)" -ForegroundColor Red
}

# 显示结果摘要
Write-Host "`n" + ("="*50) -ForegroundColor Cyan
Write-Host "分析结果摘要" -ForegroundColor Cyan
Write-Host "="*50 -ForegroundColor Cyan
Write-Host "处理的PN号: $($pns -join ', ')" -ForegroundColor Yellow
Write-Host "处理文件数: $($excelFiles.Count)" -ForegroundColor Yellow
Write-Host "成功处理: $($allPNData.Count) 个文件" -ForegroundColor Green
Write-Host "输出目录: $outputDir" -ForegroundColor Yellow
Write-Host "输出文件: OTDR数据分析报告_时间戳.xlsx" -ForegroundColor Yellow

# 特殊显示PN 997的功率信息
$pn997 = $allPNData["1831781997"]
if ($pn997 -and $pn997.ReflectancePowerData) {
    Write-Host "`nPN 997 功率分析：" -ForegroundColor Cyan
    foreach ($power in $pn997.ReflectancePowerData.Keys) {
        if ($pn997.ReflectancePowerData[$power].I.Count -gt 0) {
            Write-Host "  $power功率 距离1 I列平均值: $($pn997.ReflectancePowerData[$power].I[0])" -ForegroundColor DarkGray
        }
    }
    Write-Host "  所有功率合并 距离1 I列总平均值: $($pn997.ReflectanceData_I[0])" -ForegroundColor Green
}

if ($analysisResults.Count -gt 0) {
    Write-Host "`n差异分析结果：" -ForegroundColor Green
    foreach ($analysis in $analysisResults) {
        if ($analysis.Results.Count -gt 0) {
            Write-Host "  $($analysis.Category):" -ForegroundColor Cyan
            Write-Host "    距离1平均值: $($analysis.Results[0].ClosestPoint)" -ForegroundColor Yellow
        }
    }
}

Write-Host "="*50 -ForegroundColor Cyan

# 如果可能，打开输出目录
Write-Host "`n是否要打开输出目录？(Y/N)" -ForegroundColor Green
$response = Read-Host
if ($response -eq 'Y' -or $response -eq 'y') {
    try {
        Invoke-Item $outputDir
    } catch {
        Write-Host "无法打开目录，请手动访问: $outputDir" -ForegroundColor Yellow
    }
}

Write-Host "`n按任意键退出..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
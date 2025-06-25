Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Contract Report Export Tool"
$form.Size = New-Object System.Drawing.Size(550, 550)
$form.StartPosition = "CenterScreen"

# Contract filter
$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "Enter Contract Filter (e.g., assurant)"
$label1.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($label1)

$contractBox = New-Object System.Windows.Forms.TextBox
$contractBox.Location = New-Object System.Drawing.Point(10, 40)
$contractBox.Size = New-Object System.Drawing.Size(510, 20)
$contractBox.Text = "assurant"
$form.Controls.Add($contractBox)

# Check rollup button
$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Text = "Check Rollup Status"
$checkButton.Location = New-Object System.Drawing.Point(10, 70)
$form.Controls.Add($checkButton)

# Results display (larger)
$resultBox = New-Object System.Windows.Forms.TextBox
$resultBox.Location = New-Object System.Drawing.Point(10, 100)
$resultBox.Size = New-Object System.Drawing.Size(510, 100)
$resultBox.Multiline = $true
$resultBox.ScrollBars = "Vertical"
$form.Controls.Add($resultBox)

# HP filter
$labelHP = New-Object System.Windows.Forms.Label
$labelHP.Text = "Enter HP Filter (e.g., anthem)"
$labelHP.Location = New-Object System.Drawing.Point(10, 210)
$form.Controls.Add($labelHP)

$hpBox = New-Object System.Windows.Forms.TextBox
$hpBox.Location = New-Object System.Drawing.Point(10, 230)
$hpBox.Size = New-Object System.Drawing.Size(510, 20)
$hpBox.Text = "anthem"
$form.Controls.Add($hpBox)

# Avoid Contract filter
$labelAvoid = New-Object System.Windows.Forms.Label
$labelAvoid.Text = "Avoid Contracts Like (e.g., no med)"
$labelAvoid.Location = New-Object System.Drawing.Point(10, 260)
$form.Controls.Add($labelAvoid)

$avoidBox = New-Object System.Windows.Forms.TextBox
$avoidBox.Location = New-Object System.Drawing.Point(10, 280)
$avoidBox.Size = New-Object System.Drawing.Size(510, 20)
$avoidBox.Text = "no med"
$form.Controls.Add($avoidBox)

# File path label and input
$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Text = "Enter full file path with slash at end (e.g., C:\Reports\Report\)"
$labelPath.Location = New-Object System.Drawing.Point(10, 310)
$form.Controls.Add($labelPath)

$pathBox = New-Object System.Windows.Forms.TextBox
$pathBox.Location = New-Object System.Drawing.Point(10, 330)
$pathBox.Size = New-Object System.Drawing.Size(510, 20)
$pathBox.Text = " \\WINH\ClinicalInformatics\Client_Employer\assurant\2025.04 Premier OA\"
$form.Controls.Add($pathBox)

# Export button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Run Query and Export to Excel"
$exportButton.Location = New-Object System.Drawing.Point(10, 370)
$form.Controls.Add($exportButton)

# Connection string for rollup lookup
$conn = "Server= WINHSQLDWPROD;Database=WorkBench;Integrated Security=True"

$exportButton.Add_Click({
    $contract = $contractBox.Text
    $hp = $hpBox.Text
    $avoid = $avoidBox.Text
    $folderPath = $pathBox.Text.Trim()
    $contractClean = $contract -replace '[\\/:*?"<>|]', '_'  # Filename-safe

    $filePath = Join-Path $folderPath "$contractClean Contract Details.xlsx"

    try {
        # Ensure Export-Excel works
        Import-Module ImportExcel -ErrorAction Stop

        # --- 1. PowerShell SQL Export ---
        $query = @"
IF OBJECT_ID('tempdb..#temp') IS NOT NULL DROP TABLE #temp;

SELECT * INTO #temp FROM (
    SELECT 
        contract,
        CASE WHEN ISNULL([Plan_Start_Date],'') = '' THEN CAST(First_Auth_Date AS VARCHAR(100))
             ELSE CAST([Plan_Start_Date] AS VARCHAR(100)) END AS Approx_Start,
        sand_pbm,
        [Medical PA],
        [Pharmacy PA],
        [Medical Benefit],
        [Pharmacy Benefit]
    FROM WorkBench.dbo.ContractName_Rollup_v4
    WHERE contract LIKE '%$contract%'
        AND contract LIKE '%$hp%'
        AND contract NOT LIKE '%$avoid%'
) a;

-- Return transposed version
SELECT ColumnName, Value
FROM (
    SELECT 
        CAST(contract AS varchar(255)) AS contract,
        CAST(Approx_Start AS varchar(255)) AS Approx_Start,
        CAST(sand_pbm AS varchar(255)) AS sand_pbm,
        CAST([Medical PA] AS varchar(255)) AS [Medical PA],
        CAST([Pharmacy PA] AS varchar(255)) AS [Pharmacy PA],
        CAST([Medical Benefit] AS varchar(255)) AS [Medical Benefit],
        CAST([Pharmacy Benefit] AS varchar(255)) AS [Pharmacy Benefit]
    FROM #temp
) src
UNPIVOT (
    Value FOR ColumnName IN (
        contract,
        Approx_Start,
        sand_pbm,
        [Medical PA],
        [Pharmacy PA],
        [Medical Benefit],
        [Pharmacy Benefit]
    )
) unpvt;
"@

        $connString = "Server=WINHSQLDWPROD;Database=WorkBench;Integrated Security=True"
        $results = Invoke-Sqlcmd -Query $query -ConnectionString $connString

        if ($results.Count -gt 0) {
            $results | Export-Excel -Path $filePath -WorksheetName "ContractDetails" -AutoSize
        }

        # --- 2. SSIS Export ---
        $paramContract = '/SET \Package.Variables[User::contract].Properties[Value];"' + $contract + '"'
        $paramHP       = '/SET \Package.Variables[User::hp].Properties[Value];"' + $hp + '"'
        $paramAvoid    = '/SET \Package.Variables[User::avoid].Properties[Value];"' + $avoid + '"'
        $paramPath     = '/SET \Package.Variables[User::filePath].Properties[Value];"' + $filePath + '"'

        $ssisPackagePath = "C:\Users\fpiccorelli\Downloads\DataExport\DataExport\Package.dtsx"
        $dtexecPath = "C:\Program Files\Microsoft SQL Server\160\DTS\Binn\DTExec.exe"

        $args = @(
            "/F", "`"$ssisPackagePath`"",
            $paramContract,
            $paramHP,
            $paramAvoid,
            $paramPath
        )

        $proc = Start-Process -FilePath $dtexecPath -ArgumentList $args -NoNewWindow -Wait -PassThru

        if ($proc.ExitCode -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Both exports complete: $filePath", "Success")
        } else {
            [System.Windows.Forms.MessageBox]::Show("SSIS failed (code $($proc.ExitCode))", "Error")
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Error")
    }
})
# Finalize and run the form
$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()

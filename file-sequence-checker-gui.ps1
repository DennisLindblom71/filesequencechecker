Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main window
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Search Missing Files'
$form.Size = New-Object System.Drawing.Size(500,400)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(500,400)

# Create TableLayoutPanel
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.ColumnCount = 4
$tableLayoutPanel.RowCount = 5
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))

# Create and add controls
$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Text = 'Enter the UNC path to the folder:'
$labelPath.AutoSize = $true
$labelPath.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$tableLayoutPanel.Controls.Add($labelPath, 0, 0)
$tableLayoutPanel.SetColumnSpan($labelPath, 4)

$textBoxPath = New-Object System.Windows.Forms.TextBox
$textBoxPath.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tableLayoutPanel.Controls.Add($textBoxPath, 0, 1)
$tableLayoutPanel.SetColumnSpan($textBoxPath, 3)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = 'Browse'
$buttonBrowse.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tableLayoutPanel.Controls.Add($buttonBrowse, 3, 1)

$labelExtension = New-Object System.Windows.Forms.Label
$labelExtension.Text = 'File extension (e.g. rar):'
$labelExtension.AutoSize = $true
$labelExtension.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$tableLayoutPanel.Controls.Add($labelExtension, 0, 2)

$textBoxExtension = New-Object System.Windows.Forms.TextBox
$textBoxExtension.Text = 'rar'
$textBoxExtension.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tableLayoutPanel.Controls.Add($textBoxExtension, 1, 2)

$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = 'Search'
$buttonSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tableLayoutPanel.Controls.Add($buttonSearch, 2, 2)

$textBoxResult = New-Object System.Windows.Forms.TextBox
$textBoxResult.Multiline = $true
$textBoxResult.ScrollBars = 'Vertical'
$textBoxResult.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tableLayoutPanel.Controls.Add($textBoxResult, 0, 3)
$tableLayoutPanel.SetColumnSpan($textBoxResult, 4)

$buttonSave = New-Object System.Windows.Forms.Button
$buttonSave.Text = 'Save Results'
$buttonSave.Enabled = $false
$buttonSave.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tableLayoutPanel.Controls.Add($buttonSave, 0, 4)

$form.Controls.Add($tableLayoutPanel)

# Function to select folder
$buttonBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder"
    if ($folderBrowser.ShowDialog() -eq 'OK') {
        $textBoxPath.Text = $folderBrowser.SelectedPath
    }
})

# Main function to search for missing files
$buttonSearch.Add_Click({
    $folderPath = $textBoxPath.Text
    $fileExtension = $textBoxExtension.Text
    $files = Get-ChildItem -Path $folderPath -Filter "*.$fileExtension"
    $sortedFiles = $files | Sort-Object Name
    $missingFiles = @()

    for ($i = 1; $i -lt $sortedFiles.Count; $i++) {
        $currentFileName = $sortedFiles[$i].BaseName
        $previousFileName = $sortedFiles[$i-1].BaseName
        $currentNumber = [int]($currentFileName -replace '.*\.part(\d+)$','$1')
        $previousNumber = [int]($previousFileName -replace '.*\.part(\d+)$','$1')
        
        if ($currentNumber - $previousNumber -gt 1) {
            for ($j = $previousNumber + 1; $j -lt $currentNumber; $j++) {
                $formattedNumber = $j.ToString('000')
                $newFileName = $previousFileName -replace '\d+$', $formattedNumber
                $missingFiles += "$newFileName.$fileExtension"
            }
        }
    }

    if ($missingFiles.Count -gt 0) {
        $textBoxResult.Text = "Missing files:`r`n" + ($missingFiles -join "`r`n")
        $buttonSave.Enabled = $true
    } else {
        $textBoxResult.Text = "All files in the sequence are present."
        $buttonSave.Enabled = $false
    }
})

# Function to save results
$buttonSave.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt"
    $saveFileDialog.Title = "Save Missing Files List"
    $saveFileDialog.FileName = "MissingFiles.txt"
    
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        $textBoxResult.Text | Out-File -FilePath $saveFileDialog.FileName
        [System.Windows.Forms.MessageBox]::Show("Results saved successfully.", "Save Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Show the form
$form.ShowDialog()
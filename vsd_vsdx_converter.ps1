Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$sessionID    = -join ((48..57) + (97..122) | Get-Random -Count 6 | % {[char]$_})
$conversionID = '    '

################### LOGGING FUNCTION ###############
$LogFile = '.\vsd_converter.log'
function Write-Log {
    Param ([string]$LogString)
    $Stamp      = (Get-Date).toString('yyyy-MM-dd HH:mm:ss')
    $LogMessage = "[$Stamp][$sessionID][$conversionID] $LogString"
    Add-content $LogFile -value $LogMessage
}
Write-Log 'App STARTED'

#################### MAIN WINDOW ###################
# Main window icon
$iconBase64 = 'AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC65f8AAAAAAAAAAAAAAAAAuuX/ALrl/wC65f8AuuX/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAuuX/AAAAAAAAAAAAuuX/ALrl/wC65f8AuuX/ALrl/wC65f8AuuX/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALrl/wC65f8AuuX/ALrl/wAAAAAAAAAAAAAAAAC65f8AuuX/ALrl/wC65f8AAAAAAAAAAAAAAAAAAAAAAAAAAAC65f8AuuX/ALrl/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC65f8AuuX/AAAAAAAAAAAAAAAAAAAAAAAAAAAAuuX/ALrl/wC65f8AuuX/ALrl/wAAAAAAAAAAAAAAAAAAAAAAAAAAALrl/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADWsDD/AAAAAAAAAAAAAAAAAAAAANawMP/WsDD/1rAw/9awMP/WsDD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1rAw/9awMP8AAAAAAAAAAAAAAAAAAAAAAAAAANawMP/WsDD/1rAw/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANawMP/WsDD/1rAw/wAAAAAAAAAAAAAAAAAAAADWsDD/1rAw/9awMP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1rAw/9awMP/WsDD/1rAw/9awMP/WsDD/AAAAAAAAAADWsDD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADWsDD/1rAw/9awMP/WsDD/AAAAAAAAAAAAAAAA1rAw/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAP//AADuHwAA7AcAAOHDAADj8wAA4PsAAP//AAD//wAA7wcAAOfHAADjxwAA8DcAAPh3AAD//wAA//8AAA=='
$iconBytes  = [Convert]::FromBase64String($iconBase64)
$icoStream  = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

# Main window form
$form1                 = [System.Windows.Forms.Form]::new()
$form1.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form1.Icon            = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($icoStream).GetHIcon()))
$form1.Size            = [System.Drawing.Size]::new(450, 210)
$form1.MaximizeBox     = $False
$form1.Text            = 'VSD to VSDX Converter'
$form1.FormBorderStyle = 'FixedDialog'

# Label for source directory Text box
$sourceLabel            = [System.Windows.Forms.Label]::new()
$sourceLabel.Location   = [System.Drawing.Point]::new(10, 10)
$sourceLabel.Text       = 'Source directory:'
$sourceLabel.Width      = 100
$form1.Controls.Add($sourceLabel)

# Source directory Text box
$sourceTextBox          = [System.Windows.Forms.TextBox]::new()
$sourceTextBox.Location = [System.Drawing.Point]::new(10, 33)
$sourceTextBox.Width    = 325
$form1.Controls.Add($sourceTextBox)

# Label for destination directory Text box
$destinationLabel          = [System.Windows.Forms.Label]::new()
$destinationLabel.Location = [System.Drawing.Point]::new(10, 75)
$destinationLabel.Text     = 'Destination directory (leave blank to save in place):'
$destinationLabel.Width    = 300
$form1.Controls.Add($destinationLabel)

# Destination directory Text box
$destinationTextBox          = [System.Windows.Forms.TextBox]::new()
$destinationTextBox.Location = [System.Drawing.Point]::new(10, 98)
$destinationTextBox.Width    = 325
$form1.Controls.Add($destinationTextBox)

# Browse button for source directory
$browseSourceButton          = [System.Windows.Forms.Button]::new()
$browseSourceButton.Location = [System.Drawing.Point]::new(350, 31)
$browseSourceButton.Text     = 'Browse'
$browseSourceButton.Add_Click({
    $folderDialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $result = $folderDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $sourceTextBox.Text = $folderDialog.SelectedPath
    }
})
$form1.Controls.Add($browseSourceButton)

# Browse button for destination directory
$browseDestinationButton          = [System.Windows.Forms.Button]::new()
$browseDestinationButton.Location = [System.Drawing.Point]::new(350, 96)
$browseDestinationButton.Text     = 'Browse'
$browseDestinationButton.Add_Click({
    $folderDialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $result       = $folderDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $destinationTextBox.Text = $folderDialog.SelectedPath
    }
})
$form1.Controls.Add($browseDestinationButton)

# Close button
$closeButton          = [System.Windows.Forms.Button]::new()
$closeButton.Location = [System.Drawing.Point]::new(230, 135)
$closeButton.Text     = 'Close'
$closeButton.Add_Click({
    $form1.Close()
})
$form1.Controls.Add($closeButton)

# Convert button. Include conversion requirement checks
$convertButton            = [System.Windows.Forms.Button]::new()
$convertButton.Location   = [System.Drawing.Point]::new(140, 135)
$convertButton.Text       = 'Convert'
$convertButton.Add_Click({
    $conversionID         = -join ((48..57) + (97..122) | Get-Random -Count 4 | % {[char]$_})
    $sourceDirectory      = $sourceTextBox.Text
    $destinationDirectory = $destinationTextBox.Text
    $saveToSameDir        = if ($destinationDirectory.Trim()) { $False } else { $True }

    # Perform requirement checks for converting
    if (-not ($sourceDirectory.Trim())) {
        [System.Windows.Forms.MessageBox]::Show('Source directory cannot be empty.', 'Error', 'OK', 'WARNING')
    }
    elseif (-not (Test-Path -LiteralPath $sourceDirectory -IsValid)) {
        [System.Windows.Forms.MessageBox]::Show('Invalid directory.', 'Error', 'OK', 'WARNING')
    }
    else {
        # Get all VSD files recursively from the source directory
        $vsdFiles = Get-ChildItem $sourceDirectory -Recurse -Filter '*.vsd' -Exclude '*.vsdx'

        $vsdFilesCount = ($vsdfiles | Measure-Object).Count
        if ($vsdFilesCount -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No VSD files found for conversion.', 'Error', 'OK', 'WARNING')
        }
        else {
            # Ensure the destination directory exists, create it if necessary
            if (-not ($saveToSameDir)) {
                if (-not (Test-Path -LiteralPath $sourceDirectory)) {
                    New-Item -Path $destinationDirectory -ItemType Directory -Force
                    Write-Log "Directory created: [$destinationDirectory]."
                }
            }
            # All checks OK. Start conversion
            $form2.ShowDialog()
        }
    }
})
$form1.Controls.Add($convertButton)

# Confirmation MessageBox upon exiting the program
$form1.Add_Closing({param($sender,$e)
    $result = [System.Windows.Forms.MessageBox]::Show(`
        'Are you sure you want to exit?', `
        'Close', [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $form2.Dispose()
        $form1.Dispose()
        Write-Log 'App CLOSED'
    }
    else {
        $e.Cancel = $True
    }
})

################## PROGRESS WINDOW #################
# Progress window form
$form2                 = [System.Windows.Forms.Form]::new()
$form2.Size            = [System.Drawing.Size]::new(400, 105)
$form2.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form2.Text            = 'Converting...'
$form2.FormBorderStyle = 'FixedDialog'
$form2.ShowIcon        = $False
$form2.MaximizeBox     = $False
$form2.MinimizeBox     = $False
$form2.ControlBox      = $False

# Progress label to display current file being converted
$progressLabel          = [System.Windows.Forms.Label]::new()
$progressLabel.Location = [System.Drawing.Point]::new(5, 10)
$progressLabel.Width    = 380
$form2.controls.add($progressLabel)

# Progress bar to display conversion progress
$progressBar          = [System.Windows.Forms.ProgressBar]::new()
$progressBar.Location = [System.Drawing.Point]::new(7, 33)
$progressBar.Style    = 'Continuous'
$progressBar.Width    = 370
$form2.Controls.Add($progressBar)

################ VSD to VSDX conversion ############
$form2.Add_Shown({
    $progressLabel.Text = 'Preparing to Convert...'
    $progressBar.Step   = (1/$vsdFilesCount)*100
    try {
        # Open Visio
        $visio = New-Object -ComObject Visio.InvisibleApp
    }
    catch {
        # Visio errors (not installed, crashed, etc.)
        Write-Log 'Unable to use MS Visio:'
        Write-Log $_.toString().Trim()
        [System.Windows.Forms.MessageBox]::Show('Unable to use MS Visio.', 'Error', 'OK', 'ERROR')
        $form2.Close()
    }
    Write-Log 'Conversion started'
    $progress        = 0
    $convertedFiles  = 0
    $failedToConvert = @()
    # Loop through each VSD file found
    try {
        foreach ($file in $vsdFiles) {
            $progress++
            $vsdFile  = $file.FullName
            $fileName = $file.BaseName
            $progressBar.PerformStep()
            $progressLabel.text = "Processing file: $fileName"
            if ($saveToSameDir) {
                # Write each vsdx to their original directory
                $vsdxFile = [System.IO.Path]::ChangeExtension($file.FullName, 'vsdx')
            }
            else {
                # Write all vsdx to the given destination directory
                $vsdxFile = Join-Path -Path $destinationDirectory -ChildPath ($file.Name -replace '\.vsd$', '.vsdx')
            }
            if (Test-Path -Path $vsdxFile -PathType Leaf) {
                # Skip conversion destination file already exists
                Write-Log "Skipping   file $progress of $vsdFilesCount : $fileName"
                Write-Log "Skipped conversion for [$vsdxFile] because a file with that name already exists."
                $failedToConvert += "    [DUPLICATE] $vsdFile"
            }
            else {
                try {
                    Write-Log "Processing file $progress of $vsdFilesCount : $fileName"
                    # Convert the file to VSDX using Visio COM object
                    $document = $visio.Documents.Open($vsdFile)
                    $document.SaveAs($vsdxFile)
                    $document.Close()
                    
                    Write-Log "Converted [$vsdFile] to [$vsdxFile]."
                    $convertedFiles++
                }
                catch {
                    Write-Log "Failed to convert file [$vsdFile]:"
                    Write-Log $_.toString().Trim()
                    $failedToConvert += "    [  ERROR  ] $vsdFile"
                }
            }
        }
        # Quit Visio
        $visio.Quit()

        # Notify the user
        $progressLabel.text = 'Finished'
        Write-Log 'Conversion ended.'
        Write-Log "Converted $convertedFiles out of $vsdFilesCount files found."
        [System.Windows.Forms.MessageBox]::Show("Conversion complete.`nConverted $convertedFiles out of $vsdFilesCount files found.", 'Info', 'OK', 'INFO')
        
        if ($failedToConvert.count -gt 0) {
            Write-Log "The following files were not possible to convert:"
            Add-content $LogFile -value $failedToConvert
        }
        # Close the window
        $form2.Close()
    }
    catch {
        # Unknown error
        Write-Log 'An ERROR occured.'
        Write-Log $_.toString().Trim()
        [System.Windows.Forms.MessageBox]::Show("An error occured.`nCheck log for derails.", 'Error', 'OK', 'ERROR')
		$form2.Close()
    }
})

# Launch program (main window)
[System.Windows.Forms.Application]::Run($form1)

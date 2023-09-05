Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

################### Logging function ###############
$Logfile = ".\vsd_converter.log"
function Write-Log {
	Param ([string]$LogString)
	$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
	$LogMessage = "$Stamp $LogString"
	Add-content $LogFile -value $LogMessage
}

################ VSD to VSDX conversion ############
function Convert-VSD ($sourceDirectory, $destinationDirectory) {

	$sourceDirectory = $sourceTextBox.Text
	$destinationDirectory = $destinationTextBox.Text
	$saveToSameDir = if ($destinationDirectory.Trim()) { $false } else { $true }

	if (-not ($sourceDirectory.Trim())) {
		[System.Windows.Forms.MessageBox]::Show("Source directory cannot be empty.", 'Error', 'OK', 'WARNING')
	}
	elseif (-not (Test-Path -Path $sourceDirectory -PathType Container)) {
		[System.Windows.Forms.MessageBox]::Show("Invalid directory.", 'Error', 'OK', 'WARNING')
	}
	else {
		# Ensure the destination directory exists, create it if necessary
		if (-not ($saveToSameDir)) {
			New-Item -Path $destinationDirectory -ItemType Directory -Force
			Write-Log "Directory created: [$destinationDirectory]."
		}

		# Get all VSD files recursively from the source directory
		$vsdFiles = Get-ChildItem $sourceDirectory -Recurse -Filter "*.vsd" -Exclude "*.vsdx"

		# Open Visio
		try {
			$visio = New-Object -ComObject Visio.Application -ErrorAction Stop
		}
		catch {
			Write-Log "Unable to open MS Visio."
			Write-Log $_
			[System.Windows.Forms.MessageBox]::Show("Unable to open MS Visio.", 'Error', 'OK', 'ERROR')
			Write-Log "App CLOSED"
			$form.Close()
		}
		
		# Loop through each VSD file and convert it to VSDX using Visio COM object
		$totalFiles     = 0
		$convertedFiles = 0
		foreach ($file in $vsdFiles) {
			$totalFiles++
			$vsdFile = $file.FullName
			if ($saveToSameDir) {
				# Write each vsdx to their original directory
				$vsdxFile = [System.IO.Path]::ChangeExtension($file.FullName, "vsdx")
			}
			else {
				# Write all vsdx to the given destination directory
				$vsdxFile = Join-Path -Path $destinationDirectory -ChildPath ($file.Name -replace '\.vsd$', '.vsdx')
			}

			if (Test-Path -Path $vsdxFile -PathType Leaf) {
				# Skip conversion destination file already exists
				Write-Log "Skipped conversion for [$vsdxFile] because a file with that name already exists."
			}
			else {
				# Open the VSD
				$document = $visio.Documents.Open($vsdFile)

				# Save as VSDX
				$document.SaveAs($vsdxFile)

				# Close the document
				$document.Close()

				$convertedFiles++
				Write-Log "Converted [$vsdFile] to [$vsdxFile]."
			}
		}

		# Quit Visio
		$visio.Quit()
		Write-Log "Converted $convertedFiles out of $totalFiles files found."
		[System.Windows.Forms.MessageBox]::Show("Conversion complete.", 'Info', 'OK', 'INFO')
	}
}

Write-Log "App STARTED"

################# Create Windows Form ##############
# Form icon
$iconBase64 = 'AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC65f8AAAAAAAAAAAAAAAAAuuX/ALrl/wC65f8AuuX/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAuuX/AAAAAAAAAAAAuuX/ALrl/wC65f8AuuX/ALrl/wC65f8AuuX/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALrl/wC65f8AuuX/ALrl/wAAAAAAAAAAAAAAAAC65f8AuuX/ALrl/wC65f8AAAAAAAAAAAAAAAAAAAAAAAAAAAC65f8AuuX/ALrl/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC65f8AuuX/AAAAAAAAAAAAAAAAAAAAAAAAAAAAuuX/ALrl/wC65f8AuuX/ALrl/wAAAAAAAAAAAAAAAAAAAAAAAAAAALrl/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADWsDD/AAAAAAAAAAAAAAAAAAAAANawMP/WsDD/1rAw/9awMP/WsDD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1rAw/9awMP8AAAAAAAAAAAAAAAAAAAAAAAAAANawMP/WsDD/1rAw/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANawMP/WsDD/1rAw/wAAAAAAAAAAAAAAAAAAAADWsDD/1rAw/9awMP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1rAw/9awMP/WsDD/1rAw/9awMP/WsDD/AAAAAAAAAADWsDD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADWsDD/1rAw/9awMP/WsDD/AAAAAAAAAAAAAAAA1rAw/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8AAP//AADuHwAA7AcAAOHDAADj8wAA4PsAAP//AAD//wAA7wcAAOfHAADjxwAA8DcAAPh3AAD//wAA//8AAA=='
$iconBytes  = [Convert]::FromBase64String($iconBase64)
$icoStream  = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

$form      = [System.Windows.Forms.Form]::new()
$form.Text = "VSD to VSDX Converter"
$form.Size = [System.Drawing.Size]::new(450, 200)
$form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($icoStream).GetHIcon()))

$sourceLabel          = [System.Windows.Forms.Label]::new()
$sourceLabel.Text     = "Source directory:"
$sourceLabel.Location = [System.Drawing.Point]::new(10, 10)
$sourceLabel.Width    = 100
$form.Controls.Add($sourceLabel)

$sourceTextBox          = [System.Windows.Forms.TextBox]::new()
$sourceTextBox.Location = [System.Drawing.Point]::new(10, 35)
$sourceTextBox.Width    = 320
$form.Controls.Add($sourceTextBox)

$destinationLabel          = [System.Windows.Forms.Label]::new()
$destinationLabel.Text     = "Destination directory (leave blank to save in place):"
$destinationLabel.Location = [System.Drawing.Point]::new(10, 70)
$destinationLabel.Width    = 300
$form.Controls.Add($destinationLabel)

$destinationTextBox          = [System.Windows.Forms.TextBox]::new()
$destinationTextBox.Location = [System.Drawing.Point]::new(10, 95)
$destinationTextBox.Width    = 320
$form.Controls.Add($destinationTextBox)

$browseSourceButton          = [System.Windows.Forms.Button]::new()
$browseSourceButton.Text     = "Browse"
$browseSourceButton.Location = [System.Drawing.Point]::new(340, 33)
$browseSourceButton.Add_Click({
	$folderDialog = [System.Windows.Forms.FolderBrowserDialog]::new()
	$result = $folderDialog.ShowDialog()
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		$sourceTextBox.Text = $folderDialog.SelectedPath
	}
})
$form.Controls.Add($browseSourceButton)

$browseDestinationButton          = [System.Windows.Forms.Button]::new()
$browseDestinationButton.Text     = "Browse"
$browseDestinationButton.Location = [System.Drawing.Point]::new(340, 93)
$browseDestinationButton.Add_Click({
	$folderDialog = [System.Windows.Forms.FolderBrowserDialog]::new()
	$result = $folderDialog.ShowDialog()
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		$destinationTextBox.Text = $folderDialog.SelectedPath
	}
})
$form.Controls.Add($browseDestinationButton)

$closeButton          = [System.Windows.Forms.Button]::new()
$closeButton.Text     = "Close"
$closeButton.Location = [System.Drawing.Point]::new(230, 130)
$closeButton.Add_Click({
	Write-Log "App CLOSED"
	$form.Close()
})
$form.Controls.Add($closeButton)

$convertButton          = [System.Windows.Forms.Button]::new()
$convertButton.Text     = "Convert"
$convertButton.Location = [System.Drawing.Point]::new(140, 130)
$convertButton.Add_Click({Convert-VSD $sourceTextBox.Text $destinationTextBox.Text})
$form.Controls.Add($convertButton)

# Launch the app
[System.Windows.Forms.Application]::Run($form)

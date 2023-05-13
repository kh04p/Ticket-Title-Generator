#PACKAGES FOR GUI CREATION
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#ENABLE FLAT UI STYLE
[System.Windows.Forms.Application]::EnableVisualStyles()

#CREATE BASE FORM
$form = New-Object System.Windows.Forms.Form
$form.Text = 'CARTS Title Generator'
$form.Size = New-Object System.Drawing.Size(1290,230)
$form.StartPosition = 'CenterScreen'
$form.AutoSize = $true

#PLATFORM SELECTOR
#1. Label for selector
$labelPlatform = New-Object System.Windows.Forms.Label
$labelPlatform.Location = New-Object System.Drawing.Point(10,20)
$labelPlatform.Size = New-Object System.Drawing.Size(280,30)
$labelPlatform.Text = 'Choose a platform:'
$labelPlatform.Font = ‘Segoe UI,12’
$form.Controls.Add($labelPlatform)

#2. Drop-down menu for selector
$listPlatform = New-Object System.Windows.Forms.ComboBox
$listPlatform.Location = New-Object System.Drawing.Point(10,50)
$listPlatform.Size = New-Object System.Drawing.Size(300,20)
$listPlatform.AutoSize = $true

#3. Add default business services to drop down menu
$arraySW = Get-Content -Path @("$PSScriptRoot\platforms.txt") | Sort-Object | ForEach-Object {[void] $listPlatform.Items.Add($_)}
$listPlatform.SelectedIndex = 0
$listPlatform.Font = ‘Segoe UI,12’
$listPlatform.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($listPlatform)


#4. Import list of HW and SW issues
$issuesSW = get-content -raw '.\servicesSW.txt' | ConvertFrom-StringData
$issuesHW = get-content -raw '.\servicesHW.txt' | ConvertFrom-StringData

$arraySW = @()
$arrayHW = @()

foreach ($item in $issuesSW.Keys) {
	$itemValue = $issuesSW.Item($item)
	$tempArray = $itemValue.Split(",")
	$serviceObj = [pscustomobject]@{
		Platform = $item
		Issues = $tempArray
	}
	$arraySW += $serviceObj
}

foreach ($item in $issuesHW.Keys) {
	$itemValue = $issuesHW.Item($item)
	$tempArray = $itemValue.Split(",")
	$serviceObj = [pscustomobject]@{
		Platform = $item
		Issues = $tempArray
	}
	$arrayHW += $serviceObj
}

#4. Filter Business Service list based on Platform Selector
$listPlatform_SelectedIndexChanged = {
	$listIssue.Items.Clear()
	$listIssue.Text = $null

	foreach ($service in $arrayHW) {
		if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Hardware") {
			$tempObj = $arrayHW | Where-Object {$_.Platform -eq $listPlatform.Text}
			$tempObj.Issues | ForEach-Object {[void] $listIssue.Items.Add($_)}
			$listIssue.SelectedIndex = 0
		}
	}

	foreach ($service in $arraySW) {
		if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Software") {
			$tempObj = $arraySW | Where-Object {$_.Platform -eq $listPlatform.Text}
			$tempObj.Issues | ForEach-Object {[void] $listIssue.Items.Add($_)}
			$listIssue.SelectedIndex = 0
		}
	}
}
$listPlatform.add_SelectedIndexChanged($listPlatform_SelectedIndexChanged)



#HARDWARE/SOFTWARE SELECTOR
#1. Label for selector
$labelHWSW = New-Object System.Windows.Forms.Label
$labelHWSW.Location = New-Object System.Drawing.Point(330,20)
$labelHWSW.Size = New-Object System.Drawing.Size(170,30)
$labelHWSW.Text = 'Hardware/Software:'
$labelHWSW.Font = ‘Segoe UI,12’
$form.Controls.Add($labelHWSW)

#2. Drop-down menu for selector
$listHWSW = New-Object System.Windows.Forms.ComboBox
$listHWSW.Location = New-Object System.Drawing.Point(330,50)
$listHWSW.Size = New-Object System.Drawing.Size(170,20)
$listHWSW.AutoSize = $true
$listHWSW.Font = ‘Segoe UI,12’
@("Hardware","Software") | ForEach-Object {[void] $listHWSW.Items.Add($_)}
$listHWSW.SelectedIndex = 0
$listHWSW.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($listHWSW)
$listHWSW.add_SelectedIndexChanged($listPlatform_SelectedIndexChanged)

#ISSUE SELECTOR
#1. Label for selector
$labelIssue = New-Object System.Windows.Forms.Label
$labelIssue.Location = New-Object System.Drawing.Point(520,20)
$labelIssue.Size = New-Object System.Drawing.Size(280,30)
$labelIssue.Text = 'Choose an issue:'
$labelIssue.Font = ‘Segoe UI,12’
$form.Controls.Add($labelIssue)

#2. Drop-down menu for selector
$listIssue = New-Object System.Windows.Forms.ComboBox
$listIssue.Location = New-Object System.Drawing.Point(520,50)
$listIssue.Size = New-Object System.Drawing.Size(300,20)
$listIssue.AutoSize = $true

#3. Add default business services to drop down menu
foreach ($service in $arrayHW) {
	if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Hardware") {
		$tempObj = $arrayHW | Where-Object {$_.Platform -eq $listPlatform.Text}
		$tempObj.Issues | Sort-Object | ForEach-Object {[void] $listIssue.Items.Add($_)}
		$listIssue.SelectedIndex = 0
	}
}

foreach ($service in $arraySW) {
	if ($listPlatform.Text -eq $service.Platform -and $listHWSW.Text -eq "Software") {
		$tempObj = $arraySW | Where-Object {$_.Platform -eq $listPlatform.Text}
		$tempObj.Issues | Sort-Object | ForEach-Object {[void] $listIssue.Items.Add($_)}
		$listIssue.SelectedIndex = 0
	}
}

$listIssue.Font = ‘Segoe UI,12’
$listIssue.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($listIssue)

#ERROR INPUT
#1. Label for Error text box
$labelErr = New-Object System.Windows.Forms.Label
$labelErr.Location = New-Object System.Drawing.Point(840,20)
$labelErr.Size = New-Object System.Drawing.Size(300,30)
$labelErr.Text = 'Enter an error code or issue description:'
$labelErr.Font = ‘Segoe UI,12’
$form.Controls.Add($labelErr)

#2. Error text box
$boxErr = New-Object System.Windows.Forms.TextBox
$boxErr.Location = New-Object System.Drawing.Point(840,50)
$boxErr.Size = New-Object System.Drawing.Size(300,30)
$boxErr.AutoSize = $false
$boxErr.Multiline = $false
$boxErr.Text = "  "
$boxErr.Font = ‘Segoe UI,12’
$boxErr.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$boxErr.TextAlign = HorizontalAlignment.Center
$form.Controls.Add($boxErr)

#RESULT
#1. Label for result text box
$labelRes = New-Object System.Windows.Forms.Label
$labelRes.Location = New-Object System.Drawing.Point(10,90)
$labelRes.Size = New-Object System.Drawing.Size(100,20)
$labelRes.Text = 'Generated Title:'
$labelRes.AutoSize = $true
$labelRes.Font = ‘Segoe UI,12’
$labelRes.TextAlign = [System.Drawing.ContentAlignment]::LeftCenter
$form.Controls.Add($labelRes)

#2. Result text box
$boxRes = New-Object System.Windows.Forms.TextBox
$boxRes.Location = New-Object System.Drawing.Point(135,90)
$boxRes.Size = New-Object System.Drawing.Size(1005,30)
$boxRes.AutoSize = $false
$boxRes.Multiline = $false
$boxRes.Font = ‘Segoe UI,11’
$boxRes.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$boxRes.TextAlign = HorizontalAlignment.Center
$form.Controls.Add($boxRes)

#CLIPBOARD STATUS - Indicates whether generated title has been copied to system clipboard
$labelClip = New-Object System.Windows.Forms.Label
$labelClip.Size = New-Object System.Drawing.Size(900,40)
$LabelClip.Top = ($listIssue.Height + 120)
$labelClip.Left = ($form.Width - 1280)
$LabelClip.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$LabelClip.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold+[System.Drawing.FontStyle]::Italic)
$form.Controls.Add($LabelClip)

#GENERATE BUTTON

#1. Title Generator
$generate = {
	$arrayUnfiltered = @($listPlatform.Text, $listHWSW.Text, $listIssue.Text, $boxErr.Text)
	$arrayFiltered = @()
	$result = "  "
	$emptyFlag = $false

	#a. Filter for empty fields
	foreach ($item in $arrayUnfiltered) {
		#b. If field is empty, put EMPTY! in bold and red as warning
		if ([string]::IsNullOrWhitespace($item)) {
			$emptyFlag = $true
			switch ($item) {
				$listPlatform.Text {
					$listPlatform.Text = "  EMPTY!"
					$listPlatform.ForeColor = [System.Drawing.Color]::OrangeRed
					$listPlatform.Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
				}

				$listHWSW.Text {
					$listHWSW.Text = "  EMPTY!"
					$listHWSW.ForeColor = [System.Drawing.Color]::OrangeRed
					$listHWSW.Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
				}
				$listIssue.Text {
					$listIssue.Text = "  EMPTY!"
					$listIssue.ForeColor = [System.Drawing.Color]::OrangeRed
					$listIssue.Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
				}
				$boxErr.Text {
					$boxErr.Text = "  EMPTY!"
					$boxErr.ForeColor = [System.Drawing.Color]::OrangeRed
					$boxErr.Font = [System.Drawing.Font]::new("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
				}
			}
			continue
		}

		#c. If field is NOT empty, trim whitespace and add to filtered array
		$item = $item.Trim()
		$arrayFiltered += $item
	}
	
	#d. Add all items in filtered array to result string for output
	foreach ($item in $arrayFiltered) {
		$result += "$item"

		#e. If item is not last in array, add slashes in between items
		if ($item -ne $arrayFiltered[-1]) {
			$result += " \\ "
		}
	}

	#f. Show result string in result text box
	$boxRes.Text = $result
	
	#g. Trim whitespace and send to system clipboard then display confirmation
	$result.Trim() | clip
	if ($emptyFlag -eq $true) {
		$labelClip.Text = "Copied to system clipboard!`r`nWARNING: Did you mean to leave one of the fields empty?"
		$labelClip.ForeColor = [System.Drawing.Color]::OrangeRed
	} else {
		$labelClip.Text = 'Copied to system clipboard!'
	}	
}

#2. Generate Button
$buttonGen = New-Object System.Windows.Forms.Button
$buttonGen.Size = New-Object System.Drawing.Size(100,30)
$buttonGen.Top = ($listIssue.Height + 20)
$buttonGen.Left = ($form.Width - 130)
$buttonGen.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonGen.Text = 'GENERATE'
$buttonGen.Font = ‘Segoe UI,12’
$buttonGen.ForeColor = [System.Drawing.Color]::White
$buttonGen.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonGen.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(91,181,91)
$buttonGen.BackColor = [System.Drawing.Color]::FromArgb(91,181,91)
$buttonGen.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::DarkGreen
$buttonGen.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Green
$form.Controls.Add($buttonGen)
$buttonGen.Add_Click($generate) #Adds function to generate button

#RESET BUTTON
#1. Reset all fields and their font styles
$reset = {
	$boxErr.Text = "  "
	$boxErr.Font = ‘Segoe UI,12’
	$boxErr.ForeColor = [System.Drawing.Color]::Black

	$boxRes.Text = ""
	$boxRes.Font = 'Segoe UI,12'
	$boxRes.ForeColor = [System.Drawing.Color]::Black

	$listPlatform.SelectedIndex = 0
	$listPlatform.Font = 'Segoe UI,12'
	$listPlatform.ForeColor = [System.Drawing.Color]::Black

	$listHWSW.SelectedIndex = 0
	$listHWSW.Font = 'Segoe UI,12'
	$listHWSW.ForeColor = [System.Drawing.Color]::Black

	$listIssue.SelectedIndex = 0
	$listIssue.Font = 'Segoe UI,12'
	$listIssue.ForeColor = [System.Drawing.Color]::Black

	$labelClip.Text = ""
	$labelClip.ForeColor = [System.Drawing.Color]::Black
}

#2. Reset Button
$buttonReset = New-Object System.Windows.Forms.Button
$buttonReset.Size = New-Object System.Drawing.Size(100,30)
$buttonReset.Top = ($listIssue.Height + 60)
$buttonReset.Left = ($form.Width - 130)
$buttonReset.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonReset.Text = 'RESET'
$buttonReset.Font = ‘Segoe UI,12’
$buttonReset.ForeColor = [System.Drawing.Color]::White
$buttonReset.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonReset.FlatAppearance.BorderColor = [System.Drawing.Color]::Tomato
$buttonReset.BackColor = [System.Drawing.Color]::Tomato
$buttonReset.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::Red
$buttonReset.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::OrangeRed
$form.Controls.Add($buttonReset)
$buttonReset.Add_Click($reset) #Adds function to reset button

#SHOW FORM
$Null = $form.ShowDialog()
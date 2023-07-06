Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()


#region HIDE CONSOLE WINDOW
#Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)
#endregion
#####################

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Directory Search"
$form.ClientSize = New-Object System.Drawing.Size(550, 530)
$form.AutoSize    = $true
$form.BackColor   = 'White'
$form.StartPosition = 'CenterScreen'
$form.SizeGripStyle = "Hide"

# Create controls
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Size = New-Object System.Drawing.Size(100, 20)
$label.Font = 'lucidia,10'
$label.Text = "Search Path:"
$form.Controls.Add($label)

$textboxPath = New-Object System.Windows.Forms.TextBox
$textboxPath.Location = New-Object System.Drawing.Point(120, 10)
$textboxPath.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($textboxPath)

$groupbox_1 = New-Object System.Windows.Forms.GroupBox
$groupbox_1.Location = New-Object System.Drawing.Point(10, 70)
$groupbox_1.Size = New-Object System.Drawing.Size(480, 40)
$groupbox_1.Text = "Sort by"
$form.Controls.Add($groupbox_1)

$radioButtonName = New-Object System.Windows.Forms.RadioButton
$radioButtonName.Location = New-Object System.Drawing.Point(10, 15)
$radioButtonName.Size = New-Object System.Drawing.Size(60, 20)
$radioButtonName.Text = "Name"
$radioButtonName.Checked = $true
$groupbox_1.Controls.Add($radioButtonName)

$radioButtonSize = New-Object System.Windows.Forms.RadioButton
$radioButtonSize.Location = New-Object System.Drawing.Point(80, 15)
$radioButtonSize.Size = New-Object System.Drawing.Size(50, 20)
$radioButtonSize.Text = "Size"
$groupbox_1.Controls.Add($radioButtonSize)

$radioButtonDate = New-Object System.Windows.Forms.RadioButton
$radioButtonDate.Location = New-Object System.Drawing.Point(140, 15)
$radioButtonDate.Size = New-Object System.Drawing.Size(95, 20)
$radioButtonDate.Text = "Date Modified"
$groupbox_1.Controls.Add($radioButtonDate)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowseX = $textboxPath.Right + 4
$buttonBrowseY = 10
$buttonBrowse.Location = New-Object System.Drawing.Point($buttonBrowseX, $buttonBrowseY)
$buttonBrowse.Size = New-Object System.Drawing.Size(30, 20)
$buttonBrowse.Text = "..."
$buttonBrowse.Font             = 'Microsoft Sans Serif,10'
$buttonBrowse.FlatStyle        = 'Popup'
$buttonBrowse.FlatAppearance.BorderSize = 1
$buttonBrowse.BackColor        = 'White'
$buttonBrowse.UseVisualStyleBackColor   = $false
$form.Controls.Add($buttonBrowse)

# Create a new dropdown box
$comboboxPaths = New-Object System.Windows.Forms.ComboBox
$comboboxPathsX = $textboxPath.Left
$comboboxPathsY = $textboxPath.Bottom + 2
$comboboxPaths.Location = New-Object System.Drawing.Point($comboboxPathsX, $comboboxPathsY)
$comboboxPaths.Size = New-Object System.Drawing.Size(210, 40)
$comboboxPaths.DropDownStyle = 'DropDown'
$comboboxPaths.Text = 'Select common paths..'
[void]$comboboxPaths.Items.Add("$env:USERPROFILE")
[void]$comboboxPaths.Items.Add("$env:USERPROFILE\desktop")
[void]$comboboxPaths.Items.Add("$env:USERPROFILE\documents")
[void]$comboboxPaths.Items.Add("$env:USERPROFILE\pictures")
[void]$comboboxPaths.Items.Add("${env:ProgramFiles(x86)}")
[void]$comboboxPaths.Items.Add("$env:ProgramFiles")
[void]$comboboxPaths.Items.Add("$env:APPDATA")
[void]$comboboxPaths.Items.Add("$env:localappdata")
$form.Controls.Add($comboboxPaths)

$listview = New-Object System.Windows.Forms.ListView
$listview.Location = New-Object System.Drawing.Point(10, 120)
$listview.Size = New-Object System.Drawing.Size(505, 340)
$listview.BackColor = [System.Drawing.Color]::GhostWhite
$listview.ForeColor = [System.Drawing.Color]::Black
$listview.View = [System.Windows.Forms.View]::Details
$listview.FullRowSelect = $true
$listview.GridLines = $true
$listview.AllowColumnReorder = $true
$listview.ShowItemToolTips = $true
$listView.Alignment = [System.Windows.Forms.ListViewAlignment]::Left
$listview.Columns.Add("Name", 200) | Out-Null
$listview.Columns.Add("Size (KB)", 60) | Out-Null
$listview.Columns.Add("Date Modified", 120) | Out-Null
$listview.Columns.Add("Parent Folder", 100) | Out-Null
$form.Controls.Add($listview)

# Filter box/label
$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Location = New-Object System.Drawing.Point(10, 475)
$filterLabel.Size = New-Object System.Drawing.Size(55, 30)
$filterLabel.Text = "Filter:"
$filterLabel.Font = New-Object System.Drawing.Font("Consolas", 10)
$filterLabel.AutoSize = $true
$form.Controls.Add($filterLabel)

$filtertextBox = New-Object System.Windows.Forms.TextBox
$filtertextBox.Location = New-Object System.Drawing.Point(75, 475)
$filtertextBox.Size = New-Object System.Drawing.Size(150, 20)
$filtertextBox.Text = 'Type to filter...'
$filtertextBox.ForeColor = "Gray"
$form.Controls.Add($filtertextBox)
################

# Count and mem usage label
$countLabel = New-Object System.Windows.Forms.Label
$countLabel.Location = New-Object System.Drawing.Point(235, 475)
$countLabel.Font = New-Object System.Drawing.Font("Consolas", 12)
$countLabel.AutoSize = $true
$countLabel.ForeColor = [System.Drawing.Color]::DarkGreen
$form.Controls.Add($countLabel)
################

# NEXT button
$buttonNextItem = New-Object System.Windows.Forms.Button
$buttonNextItem.Location = New-Object System.Drawing.Point(75, 493)
$buttonNextItem.Size = New-Object System.Drawing.Size(60, 20)
$buttonNextItem.Visible = $false
$buttonNextItem.Width = 80
$buttonNextItem.Text = "Next"
$buttonNextItem.Font             = 'Microsoft Sans Serif,10'
$buttonNextItem.FlatStyle        = 'Popup'
$buttonNextItem.FlatAppearance.BorderSize = 1
$buttonNextItem.BackColor        = '#FFAFDAFF'
$buttonNextItem.UseVisualStyleBackColor   = $false
$form.Controls.Add($buttonNextItem)
################

# REFRESH BUTTON
$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Location = New-Object System.Drawing.Point(450, 475)
$buttonRefresh.Size = New-Object System.Drawing.Size(65, 20)
$buttonRefresh.Text = "Refresh"
$buttonRefresh.Font             = 'Microsoft Sans Serif,10'
$buttonRefresh.FlatStyle        = 'Popup'
$buttonRefresh.FlatAppearance.BorderSize = 1
$buttonRefresh.BackColor        = '#FFAFDAFF'
$buttonRefresh.UseVisualStyleBackColor   = $false
$form.Controls.Add($buttonRefresh)
################


################################
#       EVENTs / FUNCTIONS
################################

# BROWSER BUTTON
$buttonBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.ShowNewFolderButton = $false

    # Show the FolderBrowserDialog and get the selected folder path
    if ($dialog.ShowDialog() -eq "OK") {
        $selectedPath = $dialog.SelectedPath
        $textboxPath.Text = $selectedPath
        $textboxPath.Focus()
    }
})
################

# COMBO BOX PATHS
$comboboxPaths.add_SelectedIndexChanged({
    $textboxPath.Text = $comboboxPaths.SelectedItem
    $textboxPath.Focus()
})
################

# SEARCH BUTTON
$largeDirs = @("C:\", "C:\users",$env:HOMEDRIVE,$env:APPDATA,$env:LOCALAPPDATA,`
    $env:USERPROFILE, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:SystemRoot, $env:TEMP, $env)

$textboxPath.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        if ($largeDirs -contains $textboxPath.Text) {
            $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to search [$($textboxPath.Text)]?`r`nThis can take several minutes.", 'Confirmation', 'YesNo', 'Question')
            if ($confirmation -eq 'Yes') {
                Search-Path
            }
            else {
                return
            }
        } else {Search-Path}
        
    }
})
################

# RREFRESH BUTTON
$buttonRefresh.Add_Click({
        Search-Path
})
################

# SEARCH FUNCTION
function Search-Path {
    $sourcePath = $textboxPath.Text
    $count = 0
    if ((&{if ((split-path $sourcePath -leaf) -like "*.*") {Test-Path (split-path $sourcePath)} else {Test-Path $sourcePath}})) {

        $listview.Items.Clear()
        $listview.BeginUpdate() 
        $sorting = Get-Sorting

        Get-ChildItem $sourcePath -Recurse | Where-Object {!$_.PSIsContainer} | Sort-Object -Property $sorting -Descending | ForEach-Object {
            # CREATE LIST VIEW ITEMS
            #Item 0 - Name of file
            $item = New-Object System.Windows.Forms.ListViewItem($_.Name)
            
            # Save extra info to has table in Tag prop.
            $item.Tag = @{
                FullName = $_.FullName
                Parent = (Split-Path $_.FullName -Parent)
            }
            # Set tooltip to Full Name of path from Tag
            $item.ToolTipText = $item.Tag.FullName

            #$itemParent = Split-Path $item.Tag -Parent
            #If the file path is not = to source path Find file top lvl dir up to source, else isolate parent of source
            if ($item.Tag.Parent -ne $sourcePath) {
                $topLvlDir = (Compare-Object $sourcePath.Split("\") ($item.Tag.Parent).Split("\")).InputObject | Select-Object -First 1 
            } else {
                $topLvlDir = (Split-Path -Path $sourcePath -Leaf)
            }
            #item 1 - FIle size. Update row color is over 500Mb
            $item.SubItems.Add("{0:N0}" -f ($_.Length / 1KB))
            if ($_.Length -gt (500 * 1MB)) {$item.SubItems[1].BackColor = [System.Drawing.Color]::LightSalmon}
            #Item 2 - Date Modified
            $item.SubItems.Add("$($_.LastWriteTime)")
            if ($_.LastWriteTime -lt (Get-Date).AddYears(-1)) {$item.SubItems[2].BackColor = [System.Drawing.Color]::LightGoldenrodYellow}
            #Item 3 - Top lvl dir. up to the source dir.
            $item.SubItems.Add($topLvlDir)
            # Ensure colors affect individual subitems
            $item.UseItemStyleForSubItems = $false
            $listview.Items.Add($item)
            
            #Update file counter
            $count += 1
            $countLabel.Text = "$count files"
            $countLabel.ForeColor = [System.Drawing.Color]::DarkGreen
            $countLabel.Refresh()
        }

        $listview.EndUpdate()
    
    } else {
        # Show red message if parent directory doesn't exist and do not continue the function
        if ($count -eq 0) {
            $countLabel.Text = "No directory"
            $countLabel.ForeColor = [System.Drawing.Color]::Red
            return
        }
    }

    # Show red message if file exts don't exist (regardless of parent status)
    $fileExt = split-path $sourcePath -leaf
        if ($count -eq 0) {
            $countLabel.Text = "No $fileExt files found"
            $countLabel.ForeColor = [System.Drawing.Color]::Red
        }

    $filtertextBox.Focus()
    $countLabel.Text = "$($listview.Items.Count) files"
}
################

$listview.Add_DoubleClick({
    $itemFullName = $listview.SelectedItems[0].Tag.FullName
    if (Test-Path $itemFullName) {
        explorer /select,$itemFullName
    }
})

# FILTER
$filtertextBox.add_Enter({
    if ($filtertextBox.Text.Length -gt 1) {

        $filtertextBox.Clear()
        $filtertextBox.ForeColor = "Black"
    }
})
$filtertextBox.add_Leave({
    if ($filtertextBox.Text -eq "") {
        $filtertextBox.Text = "Type to filter..."
        $filtertextBox.ForeColor = "Gray"
    }
})

$listView.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $NextIndex = -1

    if ($listView.SelectedItems.Count -gt 0) {
        $listView.SelectedItems[0].Selected = $false
    }

    for ($i = $script:CurrentIndex + 1; $i -lt $ListView.Items.Count; $i++) {
        if ($ListView.Items[$i].BackColor -eq [System.Drawing.Color]::LightBlue) {
            $NextIndex = $i
            break;
        }
    }
    
    if ($NextIndex -eq -1) {
        for ($i = 0; $i -le $script:CurrentIndex; $i++) {
            if ($ListView.Items[$i].BackColor -eq [System.Drawing.Color]::LightBlue) {
                $NextIndex = $i
                break;
            }
        }
    }
    
    if ($NextIndex -ne -1) {
        $script:CurrentIndex = $NextIndex
        $ListView.Items[$script:CurrentIndex].Selected = $true
        $listView.Focus()
        $ListView.EnsureVisible($script:CurrentIndex)
    }
    }
})

$filtertextBox.Add_TextChanged({
    # Filter the listview control based on the text in the textbox
    $count = 0
    $buttonNextItem.Visible = $true
    if ($filtertextBox.Text -ne "Type to filter...") {
        foreach ($item in $listView.Items) {
            if ($item.Text -like "*$($filtertextBox.Text)*" ) {
                $item.ForeColor = [System.Drawing.Color]::Black
                $item.BackColor = [System.Drawing.Color]::LightBlue
                $count += 1
            } else {
                $item.ForeColor = [System.Drawing.Color]::LightGray
                $item.BackColor = [System.Drawing.Color]::White
            }
            if ($filtertextBox.Text -eq "") {
                $item.BackColor = [System.Drawing.Color]::White
                $buttonNextItem.Visible = $false
            }
        }
    }

    # Update the "found" and "total memory" labels with the number of rows found and total memory respectively
    $countLabel.Text = "$count files"
})

# Define a variable to keep track of the current selected index

# event for the NEXT item
$buttonNextItem.Add_Click({
    $NextIndex = -1

    if ($listView.SelectedItems.Count -gt 0) {
        $listView.SelectedItems[0].Selected = $false
    }

    for ($i = $script:CurrentIndex + 1; $i -lt $ListView.Items.Count; $i++) {
        if ($ListView.Items[$i].BackColor -eq [System.Drawing.Color]::LightBlue) {
            $NextIndex = $i
            break;
        }
    }
    
    if ($NextIndex -eq -1) {
        for ($i = 0; $i -le $script:CurrentIndex; $i++) {
            if ($ListView.Items[$i].BackColor -eq [System.Drawing.Color]::LightBlue) {
                $NextIndex = $i
                break;
            }
        }
    }
    
    if ($NextIndex -ne -1) {
        $script:CurrentIndex = $NextIndex
        $ListView.Items[$script:CurrentIndex].Selected = $true
        $listView.Focus()
        $ListView.EnsureVisible($script:CurrentIndex)
    }
})

function Get-Sorting {
    if ($radioButtonName.Checked) {
        return "Name"
    } elseif ($radioButtonSize.Checked) {
        return "Length"
    } elseif ($radioButtonDate.Checked) {
        return "LastWriteTime"
    }
}

# Show the form
$form.ShowDialog() | Out-Null


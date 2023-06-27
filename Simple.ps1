#*****************************************************************************************************************************
# This is the simple version of the comparison software. This file assumes that the two files entered will always be in the
# same format. Therefore, we will skip the need for a user to enter the column names in this particular version.
#
# If the formatting ever changes, check the compareName and compareGroup functions and replace "displayname",
# "mdmDisplayName", and "name" with whatever the current column headers are. These are found on lines 52, 53, 92, 93, and 98
#
#*****************************************************************************************************************************

function SelectFile([string] $initialDirectory){

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Excel files (*.csv;*.xlsx)|*.csv;*.xlsx"
    $result = $OpenFileDialog.ShowDialog()

    if($result -ne [System.Windows.Forms.DialogResult]::Cancel){

        $selectedFile.text          = $OpenFileDialog.filename
        $newFilePath                = Split-Path -Path $selectedfile.text
        $selectedFilePath.text      = Split-Path -Path $selectedfile.text
        $ReportLocation.text        = "The missing computer(s) will be written to " + $newFilePath + "\Report.txt"
        $textBoxLabel.Visible       = $true
        $textBox.Visible            = $true
        $textBoxBtn.Visible         = $true
        $fileEntry2Label.Visible    = $true
        $SelectFile2Btn.Visible     = $true
    }
}

function SelectFile2([string] $initialDirectory){

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = Split-Path -Path $selectedfile.text -Leaf -Resolve
    $OpenFileDialog.filter = "Excel files (*.csv;*.xlsx)|*.csv;*.xlsx"
    $result = $OpenFileDialog.ShowDialog()

    if($result -ne [System.Windows.Forms.DialogResult]::Cancel){
        $SelectedFile2.text         = $OpenFileDialog.FileName
        $GoBtn.Visible              = $true
    }
}

#************************************************************************************
#compareName() will read the one name and search the selected file for it.
#If the name is not found, it will be written to report.txt in the same directory as
#the file to search in.
#************************************************************************************
function compareName{

    #open the database file
    $searchData = Import-Csv -Path $selectedFile.text | select -ExpandProperty "displayname"
    $intuneMarking = Import-Csv -Path $selectedFile.text | select -ExpandProperty "mdmDisplayName" #array parallel to $searchData

    $tempPath = [IO.Path]::Combine($selectedFilePath.Text, 'Report.txt')   

    $temp = $textBox.Lines
    $found = $false

    $done.Text = "Done! " + $temp + " was found and is in Intune!"

    for(($i = 0) ; ($i -lt $searchData.Count) ; ($i++)){    
        if ([string]$temp -eq $searchData[$i]) {
            $found = $true
           
            if($intuneMarking[$i] -ne "Microsoft Intune"){
                $writeTemp = [string]$temp + " was not in Intune."
                $writeTemp | Out-File -Append $tempPath
                $done.Text = "Done! Computer found, but is not in Intune."
            }

            break
        }
    }

    if($found -ne $true){
        $writeTemp = [string]$temp + " was not found."
        $writeTemp | Out-File -Append $tempPath
        $done.Text = "Done! Computer was not found in the document."
    }

    $done.Visible = $true
}

#************************************************************************************
#compareGroup will function similarly to compareName, but will loop through a file of
#data, comparing each entry in one file to the other.
#************************************************************************************
function CompareGroup{

    #open the file to look through
    $database = Import-Csv -Path $selectedFile.text | select -ExpandProperty "displayname"
    $intuneMarking = Import-Csv -Path $selectedFile.text | select -ExpandProperty "mdmDisplayName" #array parallel to $database

    $tempPath = [IO.Path]::Combine($selectedFilePath.Text, 'Report.txt')   

    #open the file of names to search
    $searchData = Import-Csv -Path $SelectedFile2.text | select -ExpandProperty "name"

    $exceptionList

    $pbar1.Maximum = $searchData.Count
    $pBar1.Visible = $true;

    #for each member of the search file, search through the database file
    #if it exists, mark $found to be true. If not, write the result to $missing
    #$missing will be appended to the report file at the end of the function
    for(($i = 0) ; ($i -lt $searchData.Count) ; ($i++)){

        $found = $false

        for(($x = 0) ; $x -lt $database.Count ; ($x++)){
            if($database[$x] -eq $searchData[$i]){
                $found = $true

                if($intuneMarking[$x] -ne "Microsoft Intune"){
                    $writeTemp = $searchData[$i] + " was not in Intune.`n"
                    $exceptionList += $writeTemp
                }

                break
            }
        }#inner for loop

        if ($found -ne $true) {
            $exceptionList += $searchData[$i] + " was not found.`n"
        }

        $pBar1.PerformStep();
    }#outer for loop

    #write the exception list to a file
    $exceptionList | Out-File -Append $tempPath

    $done.text = "Done!"
    $done.Visible = $true
}

# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms

# Create a new form
$ComputerTestForm                    = New-Object system.Windows.Forms.Form

# Define the size, title and background color
$ComputerTestForm.AutoSize           = $true
$ComputerTestForm.text               = "Computer Search Tool"
$ComputerTestForm.BackColor          = "#ffffff"
$ComputerTestForm.MaximizeBox        = $false
$ComputerTestForm.FormBorderStyle    = 'Fixed3D'

# Create a Title for our form. We will use a label for it.
$Title                           = New-Object system.Windows.Forms.Label

# The content of the label
$Title.text                      = "Computer Search Tool"

# Make sure the label is sized the height and length of the content
$Title.AutoSize                  = $true

# Define the minial width and height (not nessary with autosize true)
$Title.width                     = 25
$Title.height                    = 10

# Position the element
$Title.location                  = New-Object System.Drawing.Point(20,20)

# Define the font type and size
$Title.Font                      = 'Microsoft Sans Serif,13'

# Other elements
$Description                     = New-Object system.Windows.Forms.Label
$Description.text                = "Select a file to scan for computer names."
$Description.AutoSize            = $false
$Description.width               = 450
$Description.height              = 20
$Description.location            = New-Object System.Drawing.Point(20,50)
$Description.Font                = 'Microsoft Sans Serif,10'

$SelectedFileLabel               = New-Object system.Windows.Forms.Label
$SelectedFileLabel.text          = "Selected File:"
$SelectedFileLabel.AutoSize      = $false
$SelectedFileLabel.width         = 450
$SelectedFileLabel.height        = 20
$SelectedFileLabel.location      = New-Object System.Drawing.Point(20,70)
$SelectedFileLabel.Font          = 'Microsoft Sans Serif,10'

$SelectedFile                    = New-Object system.Windows.Forms.Label
$SelectedFile.text               = ""
$SelectedFile.AutoSize           = $false
$SelectedFile.width              = 450
$SelectedFile.height             = 30
$SelectedFile.location           = New-Object System.Drawing.Point(20,90)
$SelectedFile.Font               = 'Microsoft Sans Serif,10'

$SelectedFilePath                = New-Object system.Windows.Forms.Label
$SelectedFilePath.text           = ""

#ReportLocation
$ReportLocation                  = New-Object system.Windows.Forms.Label
$ReportLocation.text             = ""
$ReportLocation.AutoSize         = $false
$ReportLocation.width            = 450
$ReportLocation.height           = 30
$ReportLocation.location         = New-Object System.Drawing.Point(20,120)
$ReportLocation.Font             = 'Microsoft Sans Serif,10'

#box to retrieve specific computer name
$textBoxLabel                    = New-Object System.Windows.Forms.Label
$textBoxLabel.Text               = "Please enter the name of the computer to search for here: "
$textBoxLabel.AutoSize           = $false
$textBoxLabel.width              = 450
$textBoxLabel.height             = 20
$textBoxLabel.location           = New-Object System.Drawing.Point(20,150)
$textBoxLabel.Font               = 'Microsoft Sans Serif,10'   
$textBoxLabel.Visible            = $false

$textBox                         = New-Object System.Windows.Forms.TextBox
$textBox.Location                = New-Object System.Drawing.Point(20, 170)
$textBox.Size                    = New-Object System.Drawing.Size(260,20)
$textBox.Visible                 = $false

$textBoxBtn                      = New-Object system.Windows.Forms.Button
$textBoxBtn.BackColor            = "#3FC8C4"
$textBoxBtn.text                 = "Submit Name"
$textBoxBtn.width                = 100
$textBoxBtn.height               = 20
$textBoxBtn.location             = New-Object System.Drawing.Point(290,170)
$textBoxBtn.Font                 = 'Microsoft Sans Serif,10'
$textBoxBtn.ForeColor            = "#000000"
$textBoxBtn.Visible              = $false
$textBoxBtn.Add_Click({ CompareName })

#Option to select a text file
$fileEntry2Label                 = New-Object system.Windows.Forms.Label
$fileEntry2Label.text            = "Or, select a .csv file with all computer names in one column:"
$fileEntry2Label.AutoSize        = $false
$fileEntry2Label.width           = 450
$fileEntry2Label.height          = 20
$fileEntry2Label.location        = New-Object System.Drawing.Point(20,200)
$fileEntry2Label.Font            = 'Microsoft Sans Serif,10'
$fileEntry2Label.Visible         = $false

#select File2 button
$SelectFile2Btn                  = New-Object system.Windows.Forms.Button
$SelectFile2Btn.BackColor        = "#a4ba67"
$SelectFile2Btn.text             = "Select File"
$SelectFile2Btn.width            = 90
$SelectFile2Btn.height           = 30
$SelectFile2Btn.location         = New-Object System.Drawing.Point(20,220)
$SelectFile2Btn.Font             = 'Microsoft Sans Serif,10'
$SelectFile2Btn.ForeColor        = "#ffffff"
$SelectFile2Btn.Visible          = $false
$SelectFile2Btn.Add_Click({ SelectFile2 })

#SelectedFile2 display
$SelectedFile2                   = New-Object system.Windows.Forms.Label
$SelectedFile2.text              = ""
$SelectedFile2.AutoSize          = $false
$SelectedFile2.width             = 450
$SelectedFile2.height            = 30
$SelectedFile2.location          = New-Object System.Drawing.Point(20,260)
$SelectedFile2.Font              = 'Microsoft Sans Serif,10'

#Go button
$GoBtn                           = New-Object system.Windows.Forms.Button
$GoBtn.BackColor                 = "#3FC8C4"
$GoBtn.text                      = "GO"
$GoBtn.width                     = 90
$GoBtn.height                    = 30
$GoBtn.location                  = New-Object System.Drawing.Point(20,290)
$GoBtn.Font                      = 'Microsoft Sans Serif,10'
$GoBtn.ForeColor                 = "#ffffff"
$GoBtn.Visible                   = $false
$GoBtn.Add_Click({ CompareGroup })

#done display
$done                            = New-Object system.Windows.Forms.Label
$done.text                       = "Done!"
$done.AutoSize                   = $false
$done.width                      = 450
$done.height                     = 20
$done.location                   = New-Object System.Drawing.Point(20,330)
$done.Font                       = 'Microsoft Sans Serif,10'
$done.Visible                    = $false

#select File button
$SelectFileBtn                   = New-Object system.Windows.Forms.Button
$SelectFileBtn.BackColor         = "#a4ba67"
$SelectFileBtn.text              = "Select File"
$SelectFileBtn.width             = 90
$SelectFileBtn.height            = 30
$SelectFileBtn.location          = New-Object System.Drawing.Point(370,385)
$SelectFileBtn.Font              = 'Microsoft Sans Serif,10'
$SelectFileBtn.ForeColor         = "#ffffff"
$SelectFileBtn.Add_Click({ SelectFile })

#cancel button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Cancel"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(260,385)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$ComputerTestForm.CancelButton   = $cancelBtn
$ComputerTestForm.Controls.Add($cancelBtn)

#Progress bar for file to file comparison
$pBar1                           = New-Object System.Windows.Forms.ProgressBar
$pBar1.Minimum                   = 1
$pBar1.Maximum                   = 100 #temporary value
$pBar1.width                     = 400
$pBar1.Value                     = 1
$pBar1.Step                      = 1
$pBar1.location                  = New-Object System.Drawing.Point(20,350)
$pBar1.Visible                   = $false;

#*******************************************************************
# All Elements must be above this line.
# Add elements to the form below this line
#*******************************************************************
$ComputerTestForm.controls.AddRange(@($Title, $Description, $SelectedFileLabel, $SelectedFile, $ReportLocation))
$ComputerTestForm.controls.AddRange(@($textBoxLabel, $textBox, $textBoxBtn, $fileEntry2Label,
                                      $SelectFile2Btn, $SelectedFile2, $GoBtn, $pBar1, $done, $cancelBtn, $SelectFileBtn))
# Display the form
$result = $ComputerTestForm.ShowDialog()
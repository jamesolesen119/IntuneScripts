#*************************************************************************************************
# Comparison tool, advanced version.
# Unlike the simple version of this tool, this version asks users to enter the header of the
# csv file columns that contain the computer names to check against.
# However, 1 column is hardcoded: mdmDisplayName.
# This is by request to search for whether the computer is in Intune or not, as indicated in the
# proper csv file. If this needs to be changed, alter or remove lines 61 and 103, along with the
# code blocks that rely on them, found starting on lines 75 and 132
#
# Naturally, if the column names entered do not exist, the program will run into errors that will
# cause the program to hang and become totally unresponsive.
# While probably ill-advised, this will not be checked in this script. If the user is willing to
# utilize the more advanced features, they are assumed to be able to use it properly and
# realize typos and mistakes without intervention.
#*************************************************************************************************

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
        $columnLabel.Visible        = $true
        $colTextBox.Visible         = $true
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
        $SelectedFile2Col.Visible   = $true
        $colTextBox2.Visible        = $true
    }
}

#************************************************************************************
#compareName() will read the one name and search the selected file for it.
#If the name is not found, it will be written to report.txt in the same directory as
#the file to search in.
#************************************************************************************
function compareName{

    #open the database file

    $header = [string]$colTextBox.Lines

    $searchData = Import-Csv -Path $selectedFile.text | select -ExpandProperty $header
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

    $header = [string]$colTextBox.Lines

    $database = Import-Csv -Path $selectedFile.text | select -ExpandProperty $header
    $intuneMarking = Import-Csv -Path $selectedFile.text | select -ExpandProperty "mdmDisplayName" #array parallel to $database

    $tempPath = [IO.Path]::Combine($selectedFilePath.Text, 'Report.txt')   

    #open the file of names to search

    $header = [string]$colTextBox2.Lines

    $searchData = Import-Csv -Path $SelectedFile2.text | select -ExpandProperty $header

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

#column info
$columnLabel                     = New-Object system.Windows.Forms.Label
$columnLabel.text                = "Enter the header of the column containing the data: "
$columnLabel.AutoSize            = $false
$columnLabel.width               = 350
$columnLabel.height              = 20
$columnLabel.location            = New-Object System.Drawing.Point(20,150)
$columnLabel.Font                = 'Microsoft Sans Serif,10'
$columnLabel.Visible             = $false

#box to retrieve column header
$colTextBox                      = New-Object System.Windows.Forms.TextBox
$colTextBox.Location             = New-Object System.Drawing.Point(20, 170)
$colTextBox.Size                 = New-Object System.Drawing.Size(260,20)
$colTextBox.Visible              = $false

#box to retrieve specific computer name
$textBoxLabel                    = New-Object System.Windows.Forms.Label
$textBoxLabel.Text               = "Please enter the name of the computer to search for here: "
$textBoxLabel.AutoSize           = $false
$textBoxLabel.width              = 450
$textBoxLabel.height             = 20
$textBoxLabel.location           = New-Object System.Drawing.Point(20,200)
$textBoxLabel.Font               = 'Microsoft Sans Serif,10'   
$textBoxLabel.Visible            = $false

$textBox                         = New-Object System.Windows.Forms.TextBox
$textBox.Location                = New-Object System.Drawing.Point(20, 220)
$textBox.Size                    = New-Object System.Drawing.Size(260,20)
$textBox.Visible                 = $false
 
$textBoxBtn                      = New-Object system.Windows.Forms.Button
$textBoxBtn.BackColor            = "#3FC8C4"
$textBoxBtn.text                 = "Submit Name"
$textBoxBtn.width                = 100
$textBoxBtn.height               = 20
$textBoxBtn.location             = New-Object System.Drawing.Point(290,220)
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
$fileEntry2Label.location        = New-Object System.Drawing.Point(20,250)
$fileEntry2Label.Font            = 'Microsoft Sans Serif,10'
$fileEntry2Label.Visible         = $false

#select File2 button
$SelectFile2Btn                  = New-Object system.Windows.Forms.Button
$SelectFile2Btn.BackColor        = "#a4ba67"
$SelectFile2Btn.text             = "Select File"
$SelectFile2Btn.width            = 90
$SelectFile2Btn.height           = 30
$SelectFile2Btn.location         = New-Object System.Drawing.Point(20,270)
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
$SelectedFile2.location          = New-Object System.Drawing.Point(20,310)
$SelectedFile2.Font              = 'Microsoft Sans Serif,10'

#SelectedFile2column label
$SelectedFile2col                = New-Object system.Windows.Forms.Label
$SelectedFile2col.text           = "Enter the header of the column of names to search for: "
$SelectedFile2col.AutoSize       = $false
$SelectedFile2col.width          = 350
$SelectedFile2col.height         = 20
$SelectedFile2col.location       = New-Object System.Drawing.Point(20,340)
$SelectedFile2col.Font           = 'Microsoft Sans Serif,10'
$SelectedFile2col.Visible        = $false

#box to retrieve column header
$colTextBox2                      = New-Object System.Windows.Forms.TextBox
$colTextBox2.Location             = New-Object System.Drawing.Point(20, 360)
$colTextBox2.Size                 = New-Object System.Drawing.Size(260,20)
$colTextBox2.Visible              = $false

#Go button
$GoBtn                           = New-Object system.Windows.Forms.Button
$GoBtn.BackColor                 = "#3FC8C4"
$GoBtn.text                      = "GO"
$GoBtn.width                     = 90
$GoBtn.height                    = 30
$GoBtn.location                  = New-Object System.Drawing.Point(20,380)
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
$done.location                   = New-Object System.Drawing.Point(20,420)
$done.Font                       = 'Microsoft Sans Serif,10'
$done.Visible                    = $false

#select File button
$SelectFileBtn                   = New-Object system.Windows.Forms.Button
$SelectFileBtn.BackColor         = "#a4ba67"
$SelectFileBtn.text              = "Select File"
$SelectFileBtn.width             = 90
$SelectFileBtn.height            = 30
$SelectFileBtn.location          = New-Object System.Drawing.Point(370,475)
$SelectFileBtn.Font              = 'Microsoft Sans Serif,10'
$SelectFileBtn.ForeColor         = "#ffffff"
$SelectFileBtn.Add_Click({ SelectFile })

#cancel button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Cancel"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(260,475)
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
$pBar1.location                  = New-Object System.Drawing.Point(20,440)
$pBar1.Visible                   = $false;

#*******************************************************************
# All Elements must be above this line.
# Add elements to the form below this line
#*******************************************************************
$ComputerTestForm.controls.AddRange(@($Title, $Description, $SelectedFileLabel, $SelectedFile, $ReportLocation, $columnLabel, $colTextBox))
$ComputerTestForm.controls.AddRange(@($textBoxLabel, $textBox, $textBoxBtn, $fileEntry2Label,
                                      $SelectFile2Btn, $SelectedFile2, $SelectedFile2col, $colTextBox2, $GoBtn, $pBar1, $done, $cancelBtn, $SelectFileBtn))

# Display the form
$result = $ComputerTestForm.ShowDialog()

 

 
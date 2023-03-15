<#

.Synopsis
Creates a bootable flash drive for installing WIndows 10.
.DESCRIPTION
Provides a choice of USB drives, one of which can be made bootable. Allows you to select an iso file while the script is running.
.NOTES   
Name       : Make a bootable flash drive
Author     : Manick351
Version    : 0.1
DateCreated: 2023-03-12
DateUpdated: 2018-03-03

#>


Add-Type -assembly System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
 
    $main_form = New-Object System.Windows.Forms.Form
    $main_form.Text ='Menu'
    $main_form.Width = 800
    $main_form.Height = 600
    $main_form.AutoSize = $true
    $main_form.AutoSizeMode = "GrowAndShrink"
    $main_form.TopMost = $true 
###################################################################


#Choice USB
    $Flash = gwmi win32_diskdrive | ?{$_.interfacetype -eq "USB"} | %{gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} |  %{gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | %{$_.deviceid}
    #$FormatSystem = Get-Volume | Where-Object { $_.DriveLetter -match "$Results" } | Select-Object FileSystemType
#write-Host Current file system of USB media $FormatSystem.FileSystemType -f Green
	#$diskNumber=(Get-Partition -DriveLetter $Results).DiskNumber
	#$Style = (get-disk -Number $diskNumber).PartitionStyle

#Label
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Choice USB DISK"
    $Label.Location  = New-Object System.Drawing.Point(0,5)
    $Label.AutoSize = $true
    $main_form.Controls.Add($Label)
#ComboBox
    $ComboBox = New-Object System.Windows.Forms.ComboBox
    $ComboBox.DataSource = @($flash)| ForEach-Object {[void] $ComboBox.Items.Add($_)}
    $main_form.Controls.Add($ComboBox)
    $ComboBox.Location  = New-Object System.Drawing.Point(0,30)
    $ComboBox.add_selectedIndexChanged({
    $selected = $ComboBox.selectedItem
  
    Write-Host $selected

})


#format
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Format to:"
    $Label.Location  = New-Object System.Drawing.Point(0,60)
    $Label.AutoSize = $true
    $main_form.Controls.Add($Label)
#ComboBox
    $ComboBox1 = New-Object System.Windows.Forms.ComboBox
    $ComboBox1.DataSource = @('Fat32','NTFS')
    $ComboBox1.Location  = New-Object System.Drawing.Point(0,80)
    $main_form.Controls.Add($ComboBox1)

#Choice File System
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Convert to:"
    $Label.Location  = New-Object System.Drawing.Point(0,110)
    $Label.AutoSize = $true
    $main_form.Controls.Add($Label)
#ComboBox
    $ComboBox2 = New-Object System.Windows.Forms.ComboBox
    $ComboBox2.DataSource = @('MBR','GPT')
    $ComboBox2.Location  = New-Object System.Drawing.Point(0,130)
    $ComboBox2.selectedIndex = -1
    $main_form.Controls.Add($ComboBox2)

    $ComboBox2.add_selectedIndexChanged({
    $selected = $ComboBox2.selectedIndex
#Write-Host $selected

if ($selected -eq '1')
    {
Write-Host "GPT"
        }

else {
Write-Host "MBR"
    }

})

    $ComboBox1.add_selectedIndexChanged({
    $selected = $ComboBox1.selectedIndex
#Write-Host $selected

if ($selected -eq '1')
    {
Write-Host "NTFS"
        }

else {
Write-Host "Fat32"
    }

})

#Choice ISO button
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Select ISO file"
    $Label.Location  = New-Object System.Drawing.Point(160,10)
    $Label.AutoSize = $true
    $main_form.Controls.Add($Label)
    $button_iso = New-Object System.Windows.Forms.Button
    $button_iso.Text = 'Choose'
    $button_iso.Location = New-Object System.Drawing.Point(160,30)
    $main_form.Controls.Add($button_iso)

#Selecting a response file button
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Selecting a response file"
    $Label.Location  = New-Object System.Drawing.Point(160,60)
    $Label.AutoSize = $true
    $main_form.Controls.Add($Label)
    $button_file = New-Object System.Windows.Forms.Button
    $button_file.Text = 'Select'
    $button_file.Location = New-Object System.Drawing.Point(160,80)
    $main_form.Controls.Add($button_file)

#Start
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Launching Selected Actions"
    $Label.Location  = New-Object System.Drawing.Point(240,125)
    $Label.AutoSize = $true
    $main_form.Controls.Add($Label)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = 'Start'
    $button.Location = New-Object System.Drawing.Point(300,145)
    $main_form.Controls.Add($button)

$button_iso.Add_Click(
{

#Write-Host "ISO selection. Mounting the system image in a virtual drive"
    $Volumes = (Get-Volume).Where({$_.DriveLetter}).DriveLetter
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = [System.Windows.Forms.OpenFileDialog]::new()
    $FileBrowser.Title = 'Select ISO file'
    $FileBrowser.ShowDialog()
    $FileBrowser.FileName > $null
    $ISOFile = $FileBrowser.FileName
    Add-Type -AssemblyName PresentationCore,PresentationFramework

    $Result = Mount-DiskImage ("$ISOFile")
    $ISO = (Compare-Object -ReferenceObject $Volumes -DifferenceObject (Get-Volume).Where({$_.DriveLetter}).DriveLetter).InputObject

}
)


$button_file.Add_Click(
{

#Select file answer
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = [System.Windows.Forms.OpenFileDialog]::new()
    $FileBrowser.Title = 'Select File'
    $FileBrowser.ShowDialog()
    $FileBrowser.FileName > $null
    $ISOFile = $FileBrowser.FileName
    Add-Type -AssemblyName PresentationCore,PresentationFramework

}
)

#DONE
$main_form.ShowDialog()
Pause

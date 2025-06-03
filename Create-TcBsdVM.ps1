param(
    [Parameter(Mandatory=$true, HelpMessage="Name your VM")]
    $vmname,

    [Parameter(Mandatory=$false, HelpMessage="Select your TCBSD image")]
    $tcbsdimagefile = $null,

    [Parameter(Mandatory=$false, HelpMessage="Where do you want to store the VM")]
    $vmStoragePath = $null,

    [Parameter(Mandatory=$false, HelpMessage="Where is your VirtualBox installation?")]
    $virtualBoxPath = 'C:\Program Files\Oracle\VirtualBox'
)

# Function to show file picker dialog
function Show-FilePickerDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select Beckhoff TCBSD Image"
    $fileDialog.Filter = "Image Files (*.iso;*.img)|*.iso;*.img|ISO Files (*.iso)|*.iso|IMG Files (*.img)|*.img|All Files (*.*)|*.*"
    $fileDialog.InitialDirectory = (Get-Location).Path

    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    }
    return $null
}

# Function to show folder picker dialog
function Show-FolderPickerDialog {
    param([string]$initialDirectory)

    Add-Type -AssemblyName System.Windows.Forms
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder to store the Virtual Machine"

    if (![string]::IsNullOrEmpty($initialDirectory) -and (Test-Path $initialDirectory)) {
        $folderDialog.SelectedPath = $initialDirectory
    }

    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderDialog.SelectedPath
    }
    return $null
}

# Function to get VirtualBox default VM folder
function Get-VirtualBoxDefaultVMFolder {
    try {
        Push-Location $virtualBoxPath
        $vboxManageOutput = .\VBoxManage list systemproperties | Where-Object { $_ -match "Default machine folder:" }
        if ($vboxManageOutput) {
            $defaultFolder = ($vboxManageOutput -split "Default machine folder:\s*")[1].Trim()
            if (Test-Path $defaultFolder) {
                Pop-Location
                return $defaultFolder
            }
        }
        Pop-Location
    }
    catch {
        if (Test-Path $virtualBoxPath) { Pop-Location }
        # Fallback to common default location
    }

    return "$env:USERPROFILE\VirtualBox VMs"
}

# Use full path for working directory
$workingDirectory = (Get-Location).Path

# 1) Ensure VirtualBox is installed
if (!(Test-Path $virtualBoxPath)) {
        Write-Error "VirtualBox installation not found. Please install it or specify -virtualBoxPath."
    exit 1
}

# 2) Handle VM storage location selection
if ([string]::IsNullOrEmpty($vmStoragePath)) {
    $defaultVMFolder = Get-VirtualBoxDefaultVMFolder
    Write-Output "📁  Please select where to store the Virtual Machine..."
    Write-Output "    (Default VirtualBox location: $defaultVMFolder)"

    $selectedFolder = Show-FolderPickerDialog -initialDirectory $defaultVMFolder
    if ($null -eq $selectedFolder) {
        Write-Warning "No storage location selected. Exiting."
        exit
    }

    $vmBasePath = $selectedFolder
    Write-Output "Selected storage location: $vmBasePath"
}
else {
    if (!(Test-Path $vmStoragePath)) {
        Write-Warning "Specified VM storage path does not exist: $vmStoragePath"
        $defaultVMFolder = Get-VirtualBoxDefaultVMFolder
        Write-Output "📁  Please select where to store the Virtual Machine..."
        $selectedFolder = Show-FolderPickerDialog -initialDirectory $defaultVMFolder
        if ($null -eq $selectedFolder) {
            Write-Warning "No storage location selected. Exiting."
            exit
        }
        $vmBasePath = $selectedFolder
        Write-Output "Selected storage location: $vmBasePath"
    }
    else {
        $vmBasePath = $vmStoragePath
    }
}

# 3) Handle TCBSD image file selection
if ([string]::IsNullOrEmpty($tcbsdimagefile)) {
    Write-Output "📁  Please select your TCBSD image file..."
    $selectedFile = Show-FilePickerDialog
    if ($null -eq $selectedFile) {
        Write-Warning "No file selected. Exiting."
        exit
    }
    $imagePath = $selectedFile
    Write-Output "Selected image: $(Split-Path $imagePath -Leaf)"
}
else {
    $imagePath = Join-Path $workingDirectory $tcbsdimagefile
    if (!(Test-Path $imagePath)) {
        Write-Warning "TCBSD image not found: $tcbsdimagefile"
        Write-Output "📁  Please select your TCBSD image file..."
        $selectedFile = Show-FilePickerDialog
        if ($null -eq $selectedFile) {
            Write-Warning "No file selected. Download $tcbsdimagefile into $workingDirectory, or select a different file."
            exit
        }
        $imagePath = $selectedFile
        Write-Output "Selected image: $(Split-Path $imagePath -Leaf)"
    }
}

# 4) Create & configure the VM
Set-Location $virtualBoxPath
Write-Output "🖥️  Creating VM '$vmname' (64-bit FreeBSD)"
.\VBoxManage createvm --name $vmname --basefolder $vmBasePath --ostype FreeBSD_64 --register | Out-Null
.\VBoxManage modifyvm $vmname `
    --cpus 2 --memory 1024 --vram 128 `
    --acpi on --hpet on `
    --graphicscontroller VMSVGA `
    --firmware efi64 `
    --usb on `
    --bios-logo-display-time 5000 | Out-Null

# 5) Convert the raw .iso/.img installer into a VDI
$installerImage = "TcBSD_installer.vdi"
$runtimeImage   = "TcBSD.vhd"
$vmDir          = Join-Path $vmBasePath $vmname
$installerDisk  = Join-Path $vmDir $installerImage

Write-Output "🔄  Converting image to VDI format"
.\VBoxManage convertfromraw --format VDI $imagePath $installerDisk | Out-Host

Write-Output "📏  Resizing installer disk to 8 GiB"
.\VBoxManage modifymedium disk $installerDisk --resize 8192 | Out-Host

# 6) Create SATA controller and attach disks
Write-Output "💾  Adding SATA controller"
.\VBoxManage storagectl $vmname --name SATA --add sata --controller IntelAhci --hostiocache on --bootable on | Out-Null

Write-Output "➕  Attaching installer disk to port 1"
.\VBoxManage storageattach $vmname --storagectl SATA --port 1 --device 0 --type hdd --medium $installerDisk | Out-Null

Write-Output "➕  Creating runtime disk (16 GiB)"
$runtimeDisk = Join-Path $vmDir $runtimeImage
.\VBoxManage createmedium --filename $runtimeDisk --size 16384 --format VHD | Out-Host

Write-Output "➕  Attaching runtime disk to port 0"
.\VBoxManage storageattach $vmname --storagectl SATA --port 0 --device 0 --type hdd --medium $runtimeDisk | Out-Null

# 7) Launch the VM
$vmFile = Join-Path $vmDir "$vmname.vbox"
Start-Process $vmFile

Write-Host "✅  VM created at: $vmDir"
Write-Host "✅  Please complete the TwinCAT BSD installation inside VirtualBox."

# Return to original working directory
Set-Location $workingDirectory

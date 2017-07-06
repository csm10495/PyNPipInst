#requires –runasadministrator
#author Charles Machalow
#about A script to install Python 2.7 and Pip in an automated fashion with some picked out modules via Pip.

Param(
	[string]$HttpsProxy=$FALSE,
	[bool]$OverwritePython=$TRUE,
	[bool]$OverwritePip=$TRUE,
	[bool]$DeletePythonBackups=$FALSE,
	[bool]$AddToPath=$TRUE,
	[bool]$InstallModulesWithPip=$TRUE,
	[bool]$X86Python=$TRUE,
	[string]$PythonVersion="2.7.13"
	)

$ErrorActionPreference = "Stop" # Stop on first error

function DownloadPython ($HttpsProxy, $X86Python, $PythonVersion) {
	<#
	DownloadPython($HttpsProxy, $X86Python, $PythonVersion) - Attempts to download Python via the given proxy. Downloading x86 vs x64 depnds on the X86Python parameter.
	#>
	$WebClient = New-Object System.Net.WebClient
	$DLLocation = $env:temp + "/PythonInstaller.msi"

	if ($HttpsProxy -ne $FALSE){
		Write-Host ("HttpsProxy is set to (for Python): " + $HttpsProxy)
		$WebProxy = New-Object System.Net.WebProxy($HttpsProxy ,$TRUE)
		$WebClient.Proxy = $WebProxy
	}

	if ($X86Python -eq $TRUE) {
		# 32 bit
		Write-Host ("Downloading 32-Bit Python " + $PythonVersion + "...")
		$WebClient.DownloadFile(("https://www.python.org/ftp/python/" + $PythonVersion + "/python-" + $PythonVersion + ".msi"), $DLLocation)
	} else {
		# 64 bit
		Write-Host ("Downloading 64-Bit Python " + $PythonVersion + "...")
		$WebClient.DownloadFile(("https://www.python.org/ftp/python/" + $PythonVersion + "/python-" + $PythonVersion + ".amd64.msi"), $DLLocation)
	}
	return Test-Path $DLLocation -PathType Any # Make sure file downloaded
}

function DownloadPip ($HttpsProxy) {
	<#
	DownloadPip($HttpsProxy) - Attempts to download Pip via the given proxy.
	#>
	$WebClient = New-Object System.Net.WebClient
	$DLLocation = $env:temp + "/get-pip.py"

	if ($HttpsProxy -ne $FALSE){
		Write-Host "HttpsProxy is set to (for Pip): " + $HttpsProxy
		$WebProxy = New-Object System.Net.WebProxy($HttpsProxy ,$TRUE)
		$WebClient.Proxy = $WebProxy
	}

	Write-Host "Downloading Pip..."
	$WebClient.DownloadFile("https://bootstrap.pypa.io/get-pip.py", $DLLocation)
	Write-Host "Downloading Pip... Completed"
	return Test-Path $DLLocation -PathType Any # Make sure file downloaded
}

function InstallPython ($HttpsProxy, $X86Python, $PythonVersion) {
	<#
	InstallPython($HttpsProxy, $X86Python, $PythonVersion) - Attempts to install Python via the given proxy. Installing x86 vs x64 depnds on the X86Python parameter.
	#>
	if (DownloadPython $HttpsProxy $X86Python $PythonVersion){
		$Installer = $env:temp + "/PythonInstaller.msi"
		Write-Host "Uninstalling old Python..."
		cmd /c start /wait msiexec /x $Installer /quiet
		Write-Host "Uninstalling old Python... Completed"
		Write-Host "Installing Python..."
		Start-Process -Wait ($Installer) -ArgumentList "/quiet /qn /norestart"
		Write-Host "Installing Python... Completed"
		cmd /c "mklink C:\Python27\python2.exe C:\Python27\python.exe"
		Write-Host "Linked python2 to the newly installed Python"
		return $TRUE
	}
	return $FALSE
}

function InstallPip ($HttpsProxy) {
	<#
	InstallPip($HttpsProxy) - Attempts to install Pip via the given proxy.
	#>
	if (DownloadPip $HttpsProxy){
		Write-Host "Installing Pip..."
		Start-Process -Wait "C:\Python27\python.exe " -ArgumentList ($env:temp + "/get-pip.py") -NoNewWindow
		Write-Host "Installing Pip... Completed"
		return $TRUE
	}
	return $FALSE
}

function InstallModulesFromPip ($HttpsProxy) {
	<#
	InstallModulesFromPip($HttpsProxy) - Attempts to install modules via Pip via the given proxy.
	#>
	Write-Host "Installing Modules via Pip..."
	$arguments = "-m pip install wxPython pyserial pytest pypiwin32 pycryptodome xlrd numpy pyreadline pyinstaller psutil pyyaml mkdocs markdown-fenced-code-tabs mock colorama coverage cffi pylint pytest-cov zmq protobuf wmi"

	if ($HttpsProxy -ne $FALSE){
		$arguments = $arguments + "  --proxy=" + $HttpsProxy
	}

	Start-Process -Wait "C:\Python27\python.exe" -ArgumentList $arguments -NoNewWindow
	Write-Host "Installing Modules via Pip... Complete"
	return $TRUE
}

function BackupOldPython {
	<#
	BackupOldPython - Backs up the current Python27 folder.
	#>
	if (Test-Path "C:\Python27\" -PathType Container) {
		$ToPath = ("C:\Python27_old_" + (Get-Date).ToString().replace("/", "_").replace(" ", "_").replace(":", "_") + "\")
		Write-Host "Backing up old Python27 folder to " + $ToPath
		Move-Item "C:\Python27\" $ToPath
		return Test-Path $ToPath -PathType Any # Make sure file downloaded
	}
	Write-Host "Python27 folder not found... that is ok. Won't backup."
	return $FALSE
}

function AddPythonToPath {
	<#
	AddPythonToPath - Adds C:\Python27 to the system path
	#>
	$PYPATH = "C:\Python27"
	if (!$env:path.ToLower().contains($PYPATH.ToLower())) {
		[Environment]::SetEnvironmentVariable("PATH", ($env:path + ";" + $PYPATH), "Machine")
		Write-Host "Python27 folder has been added to PATH."
	} else {
		Write-Host "Python27 folder is already in PATH. Not adding again."
	}
}

# 'Main'
$PythonVersionSplit = $PythonVersion.split(".")
if ($PythonVersionSplit.Length -eq 3){
	$PythonMajor = [convert]::ToInt16($PythonVersionSplit[0])
	$PythonMinor = [convert]::ToInt16($PythonVersionSplit[1])
	if ($PythonMajor -eq 2 -and $PythonMinor -eq 7){
		if ($DeletePythonBackups -eq $TRUE) {
			Remove-Item c:\python27_old* -recurse
		}
		if ($OverwritePython -eq $TRUE) {
			BackupOldPython
			InstallPython $HttpsProxy $X86Python $PythonVersion
		}
		if ($OverwritePip -eq $TRUE) {
			InstallPip $HttpsProxy
		}
		if ($InstallModulesWithPip -eq $TRUE) {
			InstallModulesFromPip $HttpsProxy
		}
		if ($AddToPath -eq $TRUE) {
			AddPythonToPath
		}

		# Done!
		if ($X86Python -eq $TRUE) {
			Write-Host "Done Configuring Python x86!"
		} else {
			Write-Host "Done Configuring Python x64!"
		}
	} else {
		Write-Host "Only Python 2.7.X is supported!"
		exit 1
	}
} else {
	Write-Host "Python version given was invalid, couldn't determine major/minor version"
	exit 1
}


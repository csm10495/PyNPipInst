﻿#requires –runasadministrator
#author Charles Machalow
#about A script to install Python 2.7 and Pip in an automated fashion with some picked out modules via Pip.

Param(
	[string]$Proxy=$FALSE,
	[bool]$OverwritePython=$TRUE,
	[bool]$OverwritePip=$TRUE,
	[bool]$DeletePythonBackups=$FALSE,
	[bool]$AddToPath=$TRUE,
	[bool]$InstallModulesWithPip=$TRUE,
	[bool]$X86Python=$TRUE,
	[bool]$FallbackTryNoProxy=$FALSE, # If we fail with the proxy try, without
	[string]$PythonVersion="2.7.13"
	)

$ErrorActionPreference = "Stop" # Stop on first error

function DownloadPython ($Proxy, $X86Python, $PythonVersion) {
	<#
	DownloadPython($Proxy, $X86Python, $PythonVersion) - Attempts to download Python via the given proxy. Downloading x86 vs x64 depnds on the X86Python parameter.
	#>
	$WebClient = New-Object System.Net.WebClient
	$DLLocation = $env:temp + "/PythonInstaller.msi"

	if ($Proxy -ne $FALSE){
		Write-Host ("Proxy is set to (for Python): " + $Proxy)
		$WebProxy = New-Object System.Net.WebProxy($Proxy, $TRUE)
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

function DownloadPip ($Proxy) {
	<#
	DownloadPip($Proxy) - Attempts to download Pip via the given proxy.
	#>
	$WebClient = New-Object System.Net.WebClient
	$DLLocation = $env:temp + "/get-pip.py"

	if ($Proxy -ne $FALSE){
		Write-Host ("Proxy is set to (for Pip): " + $Proxy)
		$WebProxy = New-Object System.Net.WebProxy($Proxy ,$TRUE)
		$WebClient.Proxy = $WebProxy
	}

	Write-Host "Downloading Pip..."
	$WebClient.DownloadFile("https://bootstrap.pypa.io/get-pip.py", $DLLocation)
	Write-Host "Downloading Pip... Completed"
	return Test-Path $DLLocation -PathType Any # Make sure file downloaded
}

function InstallPython ($Proxy, $X86Python, $PythonVersion) {
	<#
	InstallPython($Proxy, $X86Python, $PythonVersion) - Attempts to install Python via the given proxy. Installing x86 vs x64 depnds on the X86Python parameter.
	#>
	if (DownloadPython $Proxy $X86Python $PythonVersion){
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

function InstallPip ($Proxy) {
	<#
	InstallPip($Proxy) - Attempts to install Pip via the given proxy.
	#>
	if (DownloadPip $Proxy){
		if ($Proxy -ne $FALSE) {
			Write-Host "Installing Pip via proxy..."
			$Arguments = ($env:temp + "/get-pip.py --proxy=" + $Proxy) 
		} else {
			Write-Host "Installing Pip..."
			$Arguments = ($env:temp + "/get-pip.py") 
		}
		$RetProcess = Start-Process -Wait "C:\Python27\python.exe " -ArgumentList $Arguments -NoNewWindow -PassThru
		if ($RetProcess.ExitCode -eq 0) {
			Write-Host "Installing Pip... Completed"
			return $TRUE
		} else {
			Write-Host "Installing Pip... Failed!"
			return $FALSE
		}
	}
	return $FALSE
}

function InstallModulesFromPip ($Proxy) {
	<#
	InstallModulesFromPip($Proxy) - Attempts to install modules via Pip via the given proxy.
	#>
	Write-Host "Installing Modules via Pip..."
	$arguments = "-m pip install wget pyserial==2.7 pytest pypiwin32 pycryptodome xlrd numpy pyreadline pyinstaller psutil pyyaml mkdocs markdown-fenced-code-tabs mock colorama coverage cffi pylint pytest-cov zmq protobuf wmi pandas nose paramiko prettytable"

	if ($Proxy -ne $FALSE){
		$arguments = $arguments + "  --proxy=" + $Proxy
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
		return Test-Path $ToPath -PathType Any # Make sure file moved
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
		$Complete = $FALSE
		# todo... allow semicolon list of proxies for this list with fallback to normal at the end
		For ($i=0; $i -le 1; $i++) {
			Try {
				if ($OverwritePython -eq $TRUE) {
					BackupOldPython
					InstallPython $Proxy $X86Python $PythonVersion
					$OverwritePython = $FALSE # don't do this again if we fail later
				}
				if ($OverwritePip -eq $TRUE) {
					if (InstallPip $Proxy -eq $FALSE) {
						throw "Pip Install Failed!"
					}
				}
				if ($InstallModulesWithPip -eq $TRUE) {
					InstallModulesFromPip $Proxy
				}
				if ($AddToPath -eq $TRUE) {
					AddPythonToPath
				}
				$Complete = $TRUE
			} Catch {
				if (($Proxy -ne $FALSE) -and ($FallbackTryNoProxy -ne $FALSE) -and ($i -lt 1)) {
					$FormatString = "Warning: {0} : {1}`n{2}`n" +
					"    + Warning CategoryInfo          : {3}`n" +
					"    + Warning FullyQualifiedErrorId : {4}`n"
					$Fields = $_.InvocationInfo.MyCommand.Name,
							  $_.ErrorDetails.Message,
							  $_.InvocationInfo.PositionMessage,
							  $_.CategoryInfo.ToString(),
							  $_.FullyQualifiedErrorId
					Write-Host ($FormatString -f $Fields) # format the exception to print it, but do not throw
					Write-Host "Because of the above, about to retry without a proxy... since we failed and FallbackTryNoProxy is set to False"
					$Proxy = $False # try without proxy
				} else {
					Write-Host "Failed on final try! About to throw (and fail the script)!"
					throw # pass up current exception as we have no way to try again (either don't have FallbackTryNoProxy or $i == 1 (means we tried 0 and 1).
				}
			}
			if ($Complete -eq $TRUE) {
				break # Done!
			}
		}

		if ($Complete -eq $FALSE) {
			Write-Host "This shouldn't be possible but Complete == False... though the loop completed... what?"
			exit 3 # not complete? something went wrong... though I thought it would have thrown above
		}

		# Done!
		if ($X86Python -eq $TRUE) {
			Write-Host "Done Configuring Python x86!"
		} else {
			Write-Host "Done Configuring Python x64!"
		}
	} else {
		Write-Host "Only Python 2.7.X is supported!"
		exit 2
	}
} else {
	Write-Host "Python version given was invalid, couldn't determine major/minor version"
	exit 1
}


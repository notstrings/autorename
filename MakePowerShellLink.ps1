Add-type -AssemblyName System.Windows.Forms

function ShowFileDialog([string] $InitialDirectory) {
    $FDlg = New-Object System.Windows.Forms.OpenFileDialog
    $FDlg.Title            = "ファイルを選んでください"
    $FDlg.InitialDirectory = $InitialDirectory
    $FDlg.Filter           = "PowerShellスクリプト(*.ps1)|*.ps1"
    $FDlg.Multiselect      = $True
    $FDlg.ShowHelp         = $True
    $null = $FDlg.ShowDialog()  
    Return $FDlg.FileNames
}

function MkPshLnk([string] $sPath) {
    $dname = [System.IO.Path]::GetDirectoryName($sPath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($sPath)
    $ppath = [System.IO.Path]::Combine($dname, $fname + ".ps1")
    $lpath = [System.IO.Path]::Combine($dname, $fname + ".lnk")
    $WSH = New-Object -ComObject WScript.Shell
    $lnk = $WSH.CreateShortCut($lpath)
    $lnk.TargetPath       = "powershell.exe"
    $lnk.IconLocation     = "powershell.exe"
    $lnk.Arguments        = "-ExecutionPolicy RemoteSigned ""$ppath"""
    $lnk.WorkingDirectory = $dname
    $null = $lnk.Save()
}

foreach($elm in (ShowFileDialog -InitialDirectory $PSScriptRoot))
{
    MkPshLnk $elm
}

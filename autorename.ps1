﻿# アセンブリ追加
Add-Type -AssemblyName "Microsoft.VisualBasic"

# 各種変数初期化
$ErrorActionPreference = "Stop"
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

###############################################################################
## Common

function InfBox([string] $ttl, [string] $msg){
    return ($Host.UI.PromptForChoice($ttl, $msg, @("OK"), 0))
}

function AskBox([string] $ttl, [string] $msg, [string[]] $opt){
    return ($Host.UI.PromptForChoice($ttl, $msg, $opt, -1))
}

function VBRep([string] $text,[string] $from,[string] $to) {
    # PowerShellでは意外に面倒な大文字小文字を区別しない文字比較
    $ret = [Microsoft.VisualBasic.Strings]::Replace($text, $from, $to, 1, -1, [Microsoft.VisualBasic.CompareMethod]::Text)
    if($null -eq $ret){ $ret = "" }
    return $ret
}

# 括弧削除
Function RemoveAllBrackets([string] $sText){
    $buff = $fname
    do {
        $buff = [regex]::Replace($buff, "\([^\(]*?\)","")
    } until (
        $buff -eq [regex]::Replace($buff, "\([^\(]*?\)","")
    )
    $buff = [regex]::Replace($buff, ".*\)", "")
    $buff = [regex]::Replace($buff, "\(.*", "")
    return $buff
}

function MoveTrush([string] $FilePath) {
    $dpath = Split-Path $FilePath -Parent
    $fpath = Split-Path $FilePath -Leaf
    $shell = new-object -comobject Shell.Application
    $shell.Namespace($dpath).ParseName($fpath).InvokeVerb("delete")
}

function MoveItemWithUniqName([string] $SrcName, [string] $DstName, [bool] $isDir) {
    $sUniq = $DstName
    $lUniq = 1
    while( (Test-Path -LiteralPath $sUniq) ) {
        if ($isDir -eq $false) {
            $dname = [System.IO.Path]::GetDirectoryName($DstName)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($DstName)
            $ename = [System.IO.Path]::GetExtension($DstName)
        }else{
            $dname = [System.IO.Path]::GetDirectoryName($DstName)
            $fname = [System.IO.Path]::GetFileName($DstName)
            $ename = ""
        }
        $sUniq = [System.IO.Path]::Combine($dname, "$fname ($lUniq)" + $ename)
        $lUniq++
    }
    $null = Move-Item -LiteralPath $SrcName -Destination $sUniq -Force
}

###############################################################################

function CleanupFName([System.IO.FileInfo] $Target) {
    CleanupNodeName $Target.FullName $Target.LastWriteTime $false
}

function CleanupDName([System.IO.DirectoryInfo] $Target) {
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        CleanupDName $elm
    }
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        CleanupFName $elm
    }
    CleanupNodeName $Target.FullName $Target.CreationTime $true
}

function CleanupNodeName([string] $TargetName, [datetime] $TargetDate, [bool] $isDir) {
    try {
        # 修正前名称
        $srcname = $TargetName
        # 修正後名称
        $dstname = $TargetName
        $dstname = RestrictText $dstname $isDir
        $dstname = RestrictDate $dstname $TargetDate $isDir
        $dstname = RestrictExt $dstname $isDir
        if ($srcname -cne $dstname) {
            $null = Write-Host "---"
            $null = Write-Host "src : $srcname"
            $null = Write-Host "dst : $dstname"
            $null = MoveItemWithUniqName $srcname $dstname $isDir
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

# ファイル・フォルダ名の正規化(文字)
function RestrictText([string] $FilePath, [bool] $isDir) {
    # パスを分解
    if ($isDir -eq $false) {
        $dname = [System.IO.Path]::GetDirectoryName($FilePath)
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $ename = [System.IO.Path]::GetExtension($FilePath)
    }else{
        $dname = [System.IO.Path]::GetDirectoryName($FilePath)
        $fname = [System.IO.Path]::GetFileName($FilePath)
        $ename = ""
    }
    # ファイル名
    $fname = [regex]::Replace($fname, "[Ａ-Ｚａ-ｚ０-９　（）［］｛｝＿]+",{ 
        param($match)
        return [Microsoft.VisualBasic.Strings]::StrConv($match, [Microsoft.VisualBasic.VbStrConv]::Narrow)
    }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
    $fname = [regex]::Replace($fname, "[ｦ-ﾟ]+",{ 
        param($match)
        return [Microsoft.VisualBasic.Strings]::StrConv($match, [Microsoft.VisualBasic.VbStrConv]::Wide)
    }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
    $fname = [Microsoft.VisualBasic.Strings]::StrConv($fname, [Microsoft.VisualBasic.VbStrConv]::Uppercase)
    # 組立
    return [System.IO.Path]::Combine($dname, $fname + $ename)
}

# ファイル・フォルダ名の正規化(日付をYYYYMMDDに)
# ・なるべく確実性の高いものだけを処理する
function RestrictDate([string] $FilePath, [datetime] $FileDate, [bool] $isDir) {
    # パスを分解
    if ($isDir -eq $false) {
        $dname = [System.IO.Path]::GetDirectoryName($FilePath)
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $ename = [System.IO.Path]::GetExtension($FilePath)
    }else{
        $dname = [System.IO.Path]::GetDirectoryName($FilePath)
        $fname = [System.IO.Path]::GetFileName($FilePath)
        $ename = ""
    }
    ## YYYY-MM-DD or YYYY.MM.DD
    $fname = [regex]::Replace($fname, "(?<![0-9]+)(19|20)(\d\d)([.-])([1-9]|0[1-9]|1[0-2])(\3)([1-9]|0[1-9]|[12][0-9]|3[01])(?![0-9]+)",{
        param($match)
        $name = $match.Value.ToUpper()
        $name = $name.Replace(".","-")
        $date = [DateTime]::ParseExact($name, "yyyy-M-d", $null) 
        if($date){ return $date.ToString("yyyyMMdd") }else{ return $match.Value }
    }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
    ## YYYY年MM月DD日
    $fname = [regex]::Replace($fname, "(?<![0-9]+)(19|20)(\d\d)年([1-9]|0[1-9]|1[0-2])月([1-9]|0[1-9]|[12][0-9]|3[01])日",{
        param($match)
        $name = $match.Value.ToUpper()
        $date = [DateTime]::ParseExact($name, "yyyy年M月d日", $null) 
        if($date){ return $date.ToString("yyyyMMdd") }else{ return $match.Value }
    }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
    ## 和暦YY-MM-DD or 和暦YY.MM.DD 
    $fname = [regex]::Replace($fname, "(令和|R|平成|H|昭和|S|明治|M|大正|T)(\d{1,2})([.-])([1-9]|0[1-9]|1[0-2])(\3)([1-9]|0[1-9]|[12][0-9]|3[01])(?![0-9]+)",{
        param($match)
        $name = $match.Value.ToUpper()
        $name = $name.Replace(".","-")
        $name = $name.Replace("R","令和")
        $name = $name.Replace("H","平成")
        $name = $name.Replace("S","昭和")
        $name = $name.Replace("M","明治")
        $name = $name.Replace("T","大正")
        $info = New-Object cultureinfo("ja-jp", $true)
        $info.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar
        $date = [DateTime]::ParseExact($name, "gy-M-d", $info) 
        if($date){ return $date.ToString("yyyyMMdd") }else{ return $match.Value }
    }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
    ## 和暦YY年MM月DD日
    $fname = [regex]::Replace($fname, "(令和|R|平成|H|昭和|S|明治|M|大正|T)(\d{1,2}|元)年([1-9]|0[1-9]|1[0-2])月([1-9]|0[1-9]|[12][0-9]|3[01])日",{
        param($match)
        $name = $match.Value.ToUpper()
        $name = $name.Replace("R","令和")
        $name = $name.Replace("H","平成")
        $name = $name.Replace("S","昭和")
        $name = $name.Replace("M","明治")
        $name = $name.Replace("T","大正")
        $info = New-Object cultureinfo("ja-jp", $true)
        $info.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar
        $date = [DateTime]::ParseExact($name, "gy年M月d日", $info) 
        if($date){ return $date.ToString("yyyyMMdd") }else{ return $match.Value }
    }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
    ## YY-MM-DD or YY.MM.DD     ※表記とタイムスタンプの関係が妥当ならリネーム
    $fname = [regex]::Replace($fname, "(?<![-.0-9a-z]+)(\d\d)([.-])(0[1-9]|1[0-2])(\2)(0[1-9]|[12][0-9]|3[01])(?![0-9]+)",{
        param($match)
        $name = $match.Value.ToUpper()
        $name = $name.Replace(".","-")
        $info = New-Object cultureinfo("ja-jp", $true)
        $info.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar
        $nameyy = ($FileDate.Year).ToString().Substring(0,2) + $name
        $dateyy = [DateTime]::ParseExact($nameyy, "yyyy-M-d", $null) 
        $namegg = $FileDate.ToString("ggg", $info) + $name
        $dategg = [DateTime]::ParseExact($namegg, "gggy-M-d", $info) 
        if( ($dateyy) -and ($FileDate.Year -eq $dateyy.Year) ){
            return $dateyy.ToString("yyyyMMdd")
        }elseif( ($dategg) -and ($FileDate.Year -eq $dategg.Year) ){
            return $dategg.ToString("yyyyMMdd")
        }else{
            return $match.Value
        }
    }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
    ## YY年MM月DD日             ※表記とタイムスタンプの関係が妥当ならリネーム
    $fname = [regex]::Replace($fname, "(?<![-.0-9a-z]+)(\d\d)年([1-9]|0[1-9]|1[0-2])月([1-9]|0[1-9]|[12][0-9]|3[01])日",{
        param($match)
        $name = $match.Value.ToUpper()
        $name = $name.Replace(".","-")
        $info = New-Object cultureinfo("ja-jp", $true)
        $info.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar
        $nameyy = ($FileDate.Year).ToString().Substring(0,2) + $name
        $dateyy = [DateTime]::ParseExact($nameyy, "yyyy年M月d日", $null) 
        $namegg = $FileDate.ToString("ggg", $info) + $name
        $dategg = [DateTime]::ParseExact($namegg, "gy年M月d日", $info) 
        if( ($dateyy) -and ($FileDate.Year -eq $dateyy.Year) ){
            return $dateyy.ToString("yyyyMMdd")
        }elseif( ($dategg) -and ($FileDate.Year -eq $dategg.Year) ){
            return $dategg.ToString("yyyyMMdd")
        }else{
            return $match.Value
        }
    }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
    # 組立
    return [System.IO.Path]::Combine($dname, $fname + $ename)
}

# その他
function RestrictExt([string] $FilePath, [bool] $isDir) {
    # パスを分解
    if ($isDir -eq $false) {
        $dname = [System.IO.Path]::GetDirectoryName($FilePath)
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $ename = [System.IO.Path]::GetExtension($FilePath)
    }else{
        $dname = [System.IO.Path]::GetDirectoryName($FilePath)
        $fname = [System.IO.Path]::GetFileName($FilePath)
        $ename = ""
    }

    # ファイル名：複数の空白を一つの空白に
    $fname = [regex]::Replace($fname, "\s+", " ")   # 複数空白
    $fname = [regex]::Replace($fname, "^\s*", "")   # 先頭空白削除
    $fname = [regex]::Replace($fname, "\s+$", "")   # 末尾空白削除

    # 拡張子：小文字
    $ename = [Microsoft.VisualBasic.Strings]::StrConv($ename, [Microsoft.VisualBasic.VbStrConv]::Narrow)
    $ename = [Microsoft.VisualBasic.Strings]::StrConv($ename, [Microsoft.VisualBasic.VbStrConv]::Lowercase)
    $ename = $ename.Trim()

    # 組立
    return [System.IO.Path]::Combine($dname, $fname + $ename)
}

###############################################################################
## Main

try {
    if ($args.Length -eq 0) {
        exit
    }
    $null = Write-Host "<<Start>>"
    ForEach ($arg in $args) {
        if( Test-Path -LiteralPath $arg ){
            if ((Get-Item $arg).PSIsContainer) {
                CleanupDName (Get-Item $arg)
            } else {
                CleanupFName (Get-Item $arg)
            }
        }
    }
    $null = Write-Host "<<End>>"
    cmd /c timeout 10
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}

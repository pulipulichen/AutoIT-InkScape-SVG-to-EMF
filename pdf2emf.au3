#include <File.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>

#pragma compile(Icon, 'icons/pdf2emf.ico')

$inkscape_path = IniRead ( @ScriptDir & "\config.ini", "config", "inkscape_path", "inkscape.exe" )
$pdftocairo_path = '"' & @ScriptDir & '\poppler\bin\pdftocairo.exe"'
$svg2emf_path = '"' & @ScriptDir & '\svg2emf.exe"'

Func StripExt($FileName)
   If StringInStr($FileName, '.') Then
	  $pos = StringInStr ($FileName, '.', 2, -1)
	  $FileName = StringTrimRight($FileName, StringLen($FileName) - $pos + 1)
   EndIf
   Return $FileName
EndFunc

Func GetFileNameNoExt($sFilePath)
 If Not IsString($sFilePath) Then
	 Return SetError(1, 0, -1)
 EndIf

 Local $FileName = StringRegExpReplace($sFilePath, "^.*\\", "")
 If StringInStr(FileGetAttrib($sFilePath), "D") = False And StringInStr($FileName, '.') Then
   $FileName = StripExt($FileName)
 EndIf

 Return $FileName
EndFunc

For $i = 1 To $CmdLine[0]
   Local $filePath = $CmdLine[$i]
   If FileExists ($filePath) == 1 Then

    ;; ..\inkscape.exe --file in.svg --export-emf out.emf
    ;Local $cmd = $inkscape_path & "--file in.svg --export-emf out.emf"
    Local $svgFilename = GetFileNameNoExt($filePath) & ".svg"
	Local $cmd = $pdftocairo_path & ' -svg "' & $filePath & '" "' & $svgFilename & '"'

	;MsgBox($MB_SYSTEMMODAL, "", $cmd)
	RunWait($cmd, '', @SW_MINIMIZE)

	Local $cmdSvg2Emf = $svg2emf_path & ' "' & StripExt($filePath) & ".svg" & '"'
	RunWait($cmdSvg2Emf, '', @SW_MINIMIZE)

	FileDelete(StripExt($filePath) & ".svg")
   EndIf
Next

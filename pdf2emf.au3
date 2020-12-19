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

Func GetDir($sFilePath)
    If Not IsString($sFilePath) Then
        Return SetError(1, 0, -1)
    EndIf

    Local $FileDir = StringRegExpReplace($sFilePath, "\\[^\\]*$", "")

    Return $FileDir
EndFunc


For $i = 1 To $CmdLine[0]
   Local $filePath = $CmdLine[$i]
   If FileExists ($filePath) == 1 Then

    FileChangeDir ( GetDir($filePath) )

    ;; ..\inkscape.exe --file in.svg --export-emf out.emf
    ;Local $cmd = $inkscape_path & "--file in.svg --export-emf out.emf"
	Local $tmpPDFFilename = GetDir($filePath) & "\tmp.pdf"
	FileCopy($filePath, $tmpPDFFilename)

    Local $finalEmfPath = GetFileNameNoExt($filePath) & ".emf"
    FileDelete($finalEmfPath)

    Local $svgFilename = GetFileNameNoExt($tmpPDFFilename) & ".svg"
 	Local $cmd = $pdftocairo_path & ' -f 1 -l 1 -svg "' & $filePath & '" "' & $svgFilename & '"'

	;MsgBox($MB_SYSTEMMODAL, "", $cmd)
	;Exit
	RunWait($cmd, '', @SW_MINIMIZE)

    FileDelete(GetFileNameNoExt($tmpPDFFilename) & ".emf")

	Local $cmdSvg2Emf = $svg2emf_path & ' "' & StripExt($tmpPDFFilename) & ".svg" & '"'
	;MsgBox($MB_SYSTEMMODAL, "", $cmdSvg2Emf)
	RunWait($cmdSvg2Emf, '', @SW_MINIMIZE)

	FileDelete(StripExt($tmpPDFFilename) & ".svg")
	FileDelete($tmpPDFFilename)

	FileMove(GetFileNameNoExt($tmpPDFFilename) & ".emf", GetFileNameNoExt($filePath) & ".emf")


	Local $cmdOpenEmf = $inkscape_path & ' "' & GetFileNameNoExt($filePath) & ".emf" & '"'
	;MsgBox($MB_SYSTEMMODAL, "", $cmdSvg2Emf)
	Run($cmdOpenEmf, '', @SW_MINIMIZE)
	;ShellExecute($inkscape_path,
   EndIf
Next

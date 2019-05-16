Option Compare Database
Option Explicit


Sub BasicCableMixedFiber()

DoCmd.OpenForm "frmFiberSpec"
DoCmd.OpenQuery "qryUpdateNewBasicCableMixedFiber"
DoCmd.Close acForm, "frmFiberSpec"


End Sub

Sub FindPath()

Dim strPath As String

strPath = Left(CurrentDb.name, InStrRev(CurrentDb.name, "\"))

MsgBox "current path name to database is " & vbCrLf & _
strPath

MsgBox "current path to msaccess.exe is " & vbCrLf & _
SysCmd(acSysCmdAccessDir)

Debug.Print "current path to msaccess.exe is " & vbCrLf & _
SysCmd(acSysCmdAccessDir)


End Sub
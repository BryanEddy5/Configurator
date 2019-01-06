Attribute VB_Name = "BasicCableMixedFiber"
Option Compare Database
Option Explicit


Function BasicCableMixedFiber()

DoCmd.OpenForm "frmFiberSpec"
DoCmd.OpenQuery "qryUpdateNewBasicCableMixedFiber"
DoCmd.Close acForm, "frmFiberSpec"


End Function

Sub FindPath()

Dim strPath As String

strPath = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\"))

MsgBox "current path name to database is " & vbCrLf & _
strPath

MsgBox "current path to msaccess.exe is " & vbCrLf & _
SysCmd(acSysCmdAccessDir)

Debug.Print "current path to msaccess.exe is " & vbCrLf & _
SysCmd(acSysCmdAccessDir)


End Sub

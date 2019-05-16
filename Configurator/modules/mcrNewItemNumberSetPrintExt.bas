Option Compare Database
Option Explicit
'------------------------------------------------------------
' mcrNewItemNumberSetPrintExt
'
'------------------------------------------------------------
Function mcrNewItemNumberSetPrintExt()
On Error GoTo mcrNewItemNumberSetPrintExt_Err

    ' set item to copy to base item & "-00"
    Forms!frmNewItemNumber!cboItemNumberCopy = Forms!frmMainMenu!ItemNumberSelect
    ' turn on New Print Extension command button
    Forms!frmNewItemNumber!cmdNewPrintExtension.Visible = True
    ' turn off Copy BOM/Routing command button
    DoCmd.SetWarnings False

    DoCmd.RunMacro "mcrItemNumberSelectCopy", , ""


mcrNewItemNumberSetPrintExt_Exit:
    Exit Function

mcrNewItemNumberSetPrintExt_Err:
    MsgBox Error$
    Resume mcrNewItemNumberSetPrintExt_Exit

End Function
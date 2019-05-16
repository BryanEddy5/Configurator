Option Compare Database
Option Explicit
'------------------------------------------------------------
' mcrNewItemNumberContinue
'
'------------------------------------------------------------
Public Function NewItemNumberContinue()
'On Error GoTo mcrNewItemNumberContinue_Err

DoCmd.SetWarnings False

    If (Eval("[Forms]![frmNewItemNumber]![NewItemNumber] Is Null")) Then
        Beep
        MsgBox "You must enter a new item number to proceed.", vbCritical, "NEW ITEM NUMBER REQUIRED"
    End If
    If (Eval("[Forms]![frmNewItemNumber]![NewItemNumber] Is Null")) Then
        End
    End If
    If (Eval("[Forms]![frmNewItemNumber]![LotSize] Is Null")) Then
        Beep
        MsgBox "You must enter a lot size to proceed.", vbCritical, "LOT SIZE REQUIRED"
    End If
    If (Eval("[Forms]![frmNewItemNumber]![LotSize] Is Null")) Then
        End
    End If

    If (Forms!frmNewItemNumber!ExistingItem = 2) Then
        End
    End If
    ' Record new base item on main menu
    Forms!frmMainMenu!NewBaseItem = Forms!frmNewItemNumber!NewItemNumber
    ' Reset print extension
    Forms!frmMainMenu!NewPrintExtension = ""
    'DoCmd.RunMacro "mcrCopyItem", , ""
    If Not CopyItem Then Exit Function
    ' Update known values
    'DoCmd.RunMacro "mcrNewItemNumberUpdate", , ""
    Call NewItemNumberUpdate

DoCmd.SetWarnings True
'mcrNewItemNumberContinue_Exit:
'    Exit Function
'
'mcrNewItemNumberContinue_Err:
'    MsgBox Error$
'    Resume mcrNewItemNumberContinue_Exit

End Function
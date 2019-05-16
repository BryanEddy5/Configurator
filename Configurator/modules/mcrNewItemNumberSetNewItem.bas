Option Compare Database
Option Explicit
'------------------------------------------------------------
' mcrNewItemNumberSetNewItem
'
'------------------------------------------------------------
Function NewItemNumberSetNewItem()
On Error GoTo mcrNewItemNumberSetNewItem_Err

    ' set item to copy to base item & "-00"
    Forms!frmNewItemNumber!cboItemNumberCopy = Null
    ' turn on Copy BOM/Routing command button
    'Forms!frmNewItemNumber!cmdCopyItem.Visible = True
    ' turn off New Print Extension command button
    'Forms!frmNewItemNumber!cmdNewPrintExtension.Visible = False
    ' disable drop-down
    Forms!frmNewItemNumber!cboItemNumberCopy.Enabled = True
    ' disable drop-down
    Forms!frmNewItemNumber!cboItemNumberCopy.Locked = False
    'DoCmd.RunMacro "mcrItemNumberSelectCopy", , ""
    [Forms]![frmMainMenu]![ItemNumberSelect] = [Forms]![frmNewItemNumber]![cboItemNumberCopy]

mcrNewItemNumberSetNewItem_Exit:
    Exit Function

mcrNewItemNumberSetNewItem_Err:
    MsgBox Error$
    Resume mcrNewItemNumberSetNewItem_Exit

End Function


Function OraclePush(ItemNumber As String)

    Dim strExt As String
    Dim CableCon As New clsCableConstruction
    Dim item As New clsCableQualities
    item.Load (ItemNumber)
    CableCon.Load (item.Base)
    Dim hold$()
    
    strExt = Mid(ItemNumber, 7, 8)
    hold = Split(ItemNumber, "-")
    

    If CurrentProject.AllForms("frmMainMenu").IsLoaded Then
    DoCmd.Close acForm, "frmMainMenu"
    End If

    DoCmd.OpenForm "frmMainMenu"
    Forms.frmMainMenu.Visible = False
    DoCmd.OpenForm "frmNewItemNumber"
    
    
    If strExt Like "*bx*" Then
        Forms.frmNewItemNumber.NewItemNumber = ItemNumber
    
    ElseIf strExt Like "*ia*" Then
        Forms.frmNewItemNumber.NewItemNumber = ItemNumber
        
    ElseIf strExt Like "*test*" Then
        Forms.frmNewItemNumber.NewItemNumber = ItemNumber
        
    ElseIf ItemNumber Like "rbn*" Or ItemNumber Like "dns*" Or ItemNumber Like "pt*" Or ItemNumber Like "pf*" Then
        Forms.frmNewItemNumber.NewItemNumber = ItemNumber
        
    ElseIf Not CableCon.IsCommoned Then
        Forms.frmNewItemNumber.NewItemNumber = ItemNumber
    Else
    
        Forms.frmNewItemNumber.NewItemNumber = hold(0) & "-00"
    
    End If
    
    
    
    
    Call Forms("frmNewItemNumber").NewItemNumber_AfterUpdate
    
    
    If IsNull(Forms!frmNewItemNumber!cboItemNumberCopy.ItemData(1)) = True Then
        MsgBox "No active items available to copy", vbCritical
    Else
    Forms!frmNewItemNumber!cboItemNumberCopy = Forms!frmNewItemNumber!cboItemNumberCopy.ItemData(1)
    End If
    
    Call Forms("frmNewItemNumber").cboItemNumberCopy_AfterUpdate


End Function
Option Compare Database
Option Explicit
Private CableConstruction As New clsCableConstruction

'------------------------------------------------------------
' mcrCreateBasicCableSheath
'
'
'Run's cable creation of 2 pass cable
'------------------------------------------------------------
Function mcrCreateBasicCableSheath()
On Error GoTo mcrCreateBasicCableSheath1_Err

DoCmd.SetWarnings False

    ' *****  AddMenu action not available
    ' Open Fiber Spec form to copy information
    DoCmd.OpenForm "frmFiberSpec", acNormal, "", "", , acNormal
    ' Clear All Tables
    DoCmd.OpenQuery "qryNewBasicCableDelete2", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNewBasicCableDelete", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNewPrductConstructionClear", acViewNormal, acEdit
    DoCmd.OpenQuery "qrytblUnitSeriesCopyClear", acViewNormal, acEdit
    
    If (Forms!frmCreateItem!Armor = 0) Then
        ' Open Special Instruction Form
        'DoCmd.OpenForm "frmSpecialPrintMatch", acNormal, "", "", , acNormal
        ' creates item
        DoCmd.OpenQuery "qryNewBasicCable_New", acViewNormal, acEdit
        ' Updates informaiton to newly created item
        DoCmd.OpenQuery "qryUpdateNewBasicCable", acViewNormal, acEdit
    End If
    
    If (Forms!frmCreateItem!Armor = -1) Then
        DoCmd.OpenQuery "qryNewArmoredCable", acViewNormal, acEdit
        ' Updates Armored data
        DoCmd.OpenQuery "qryLOADArmor_New", acViewNormal, acEdit
        ' Appends to BasicCableHold
        DoCmd.OpenQuery "qryLOADBasicCable_New", acViewNormal, acEdit
        'DoCmd.OpenQuery "qryLOADtblCableCableCopy", acViewNormal, acEdit
    End If
     
    'If (Forms!frmCreateItem!PrintType <> 0) Then
        ' Append to new Basic Cable Hold
        DoCmd.OpenQuery "qryLOADBasicCable_New", acViewNormal, acEdit
   ' End If
    
    If (Forms!frmCreateItem!Armor = -1) Then
        DoCmd.OpenForm "frmFinishNewSubunit", acNormal, "", "", , acNormal
        DoCmd.Close acForm, "frmFiberSpec"
        Exit Function
        
        
    End If
    If (Forms!frmCreateItem!Unprinted = -1) Then
        ' Creates unprinted item
        DoCmd.OpenQuery "qryNewBasicCableUnprintedUpdate", acViewNormal, acEdit
        ' appends to holding table
        DoCmd.OpenQuery "qryLOADBasicCable_New", acViewNormal, acEdit
    End If
    
    'If cable is unprinted, but a customer item then remove all print information
    If CreateNewItem.CableQualities.IsPrinted = False And Forms!frmCreateItem!PrintType <> 0 Then
        DoCmd.OpenQuery "qryNewBasicCableUnprintedUpdateCustomer"
    
    End If
    
    
    
    If (Forms!frmCreateItem![Exists] = 0) Then
        ' Finds new sub for unit series
        'DoCmd.RunMacro "mcrUnitSeries", , ""
        Call UnitSeries
    End If
    
    If (Forms!frmCreateItem![Exists] = -1) Then
        ' Update Unit Series with appropriate data
        'DoCmd.OpenQuery "qryUpdateUnitSeriesExisting", acViewNormal, acEdit
    End If
    
    ' closes fiber spec
    DoCmd.Close acForm, "frmFiberSpec"
    'DoCmd.Close acForm, "frmSpecialPrintMatch"
    DoCmd.Close acForm, "frmWait"
    DoCmd.OpenForm "frmFinishNewSubunit", acNormal, "", "", , acNormal


mcrCreateBasicCableSheath1_Exit:
    Exit Function

mcrCreateBasicCableSheath1_Err:
    MsgBox Error$
    Resume mcrCreateBasicCableSheath1_Exit

End Function


'------------------------------------------------------------
' mcrUnitSeries
'
'------------------------------------------------------------
Function UnitSeries()
On Error GoTo mcrUnitSeries_Err

CableConstruction.Load (Forms!frmCreateItem!cbSubunit)

    DoCmd.OpenForm "frmFiberMode", acNormal, "", "", , acNormal
    ' Clear Table
    'DoCmd.OpenQuery "qrytblUnitSeriesCopyClear", acViewNormal, acEdit
    ' Append to temporary Unit Series
    'DoCmd.OpenQuery "qryCableSheathUnitSeriesCopyAppend", acViewNormal, acEdit
    ' Update Unit Series to Chosen
    'DoCmd.OpenQuery "qryUpdateUnitSeries", acViewNormal, acEdit
    InsertUnitSeries (Forms!frmCreateItem!Base) ' Creates Template for unit series subunits to be added
    If (Forms!frmCreateItem!SubX = 1) Then 'And (CableConstruction.NumPositions - CableConstruction.NumFillers) > 1 Then
        ' Runs code to create sequentially colored subunits
        Call frmSubunitX
        ' Stops Macro
        Exit Function
    End If
    ' Find if the Unit Series Subunit Exists
    'DoCmd.OpenForm "frmSubUnitExists", acNormal, "", "", , acNormal
    'If (Forms!frmSubUnitExists![Exists] = -1) Then
    If CountSubs(Forms!frmCreateItem!Base) = 0 Then
        Exit Function
    End If
    If SubUnitExist() <> "0" Then
        ' Clears tlbUnitSeriesCopy
        'DoCmd.OpenQuery "qryUnitSeriesSubunitClear", acViewNormal, acEdit

        ' Appends the Unit Series to copy from
        'DoCmd.OpenQuery "qryUnitSeriesSubunitAppend", acViewNormal, acEdit
        DoCmd.OpenQuery "qryUnitSeriesSubunit_X_Append_New"

        ' Updates the Unit Series copy with new subunit
        DoCmd.OpenQuery "qryUnitSeriesCopyItemNoUpdate", acViewNormal, acEdit
        
    Else
        ' Run Macro to create new subunit
        'DoCmd.RunMacro "mcrCreateNewSubUnit", , ""
        Call mcrCreateNewSubUnit
    End If

    ' Closes form for finding if subunit exists
    DoCmd.Close acForm, "frmSubUnitExists"
    DoCmd.Close acForm, "frmFiberMode"


mcrUnitSeries_Exit:
    Exit Function

mcrUnitSeries_Err:
    MsgBox Error$
    Resume mcrUnitSeries_Exit

End Function


Function mcrCreateNewSubUnit()
On Error GoTo mcrCreateNewSubUnit_Err




    ' Open form showing Item to copy and next subunit oracle number
    'DoCmd.OpenForm "frmSubUnitNextItemNo", acNormal, "", "", , acNormal
    
    DoCmd.OpenForm "frmFiberSpec", acNormal, "", "", , acNormal
    
    ' clear table to hold subunit information
    DoCmd.OpenQuery "qryNewBasicConstructionClear", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNewPrductConstructionHoldClear", acViewNormal, acEdit
    
    ' Add item to copy's information to table
    '[Forms]![frmCreateItem]![SubColor] = Mid(Forms!frmCreateItem!ItemNo, 11, 1)
    DoCmd.OpenQuery "qrySubunitAppend", acViewNormal, acEdit
    'DoCmd.OpenForm "frmSubUnitJacketColor_X", acNormal, "", "", , acNormal
    'DoCmd.OpenForm "frmTBMaterial", acNormal, "", "", , acNormal
    DoCmd.OpenQuery "qrySubunitUpdate_X", acViewNormal, acEdit
    
    If CableConstruction.TightBuffered Then
        Call Forms("frmCreateItem").TBMaterial_Update
    Else
        DoCmd.OpenQuery "qryfrmCreateItemUpdateTBMaterial_NoMaterial"
    End If

    
    ' Update the information in the table to create new item
    
    
    ' Update the TB Material for the subunit
    'DoCmd.OpenQuery "qryfrmCreateItemTBMaterialUpdate", acViewNormal, acEdit
    
    ' Updates the tblUnitSeriesCopy with the newly created item
    DoCmd.OpenQuery "qrySubunitNewUnitSeriesUpdate", acViewNormal, acEdit
    
    ' Appends to Basic Product Hold table
    DoCmd.OpenQuery "qryNewBasicConstructionAppend", acViewNormal, acEdit
    

    'DoCmd.Close acForm, "frmTBMaterial"
    DoCmd.Close acForm, "frmFiberSpec"
    'DoCmd.Close acForm, "frmSubUnitJacketColor_X"
    DoCmd.Close acForm, "frmSubUnitNextItemNo"
    DoCmd.OpenForm "frmFinishNewSubunit", acNormal, "", "", , acNormal


mcrCreateNewSubUnit_Exit:
    Exit Function

mcrCreateNewSubUnit_Err:
    MsgBox Error$
    Resume mcrCreateNewSubUnit_Exit

End Function
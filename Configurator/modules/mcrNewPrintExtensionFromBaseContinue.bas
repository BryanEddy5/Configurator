Option Explicit
Option Compare Database

'------------------------------------------------------------
' mcrNewPrintExtensionFromBaseContinue
'
'------------------------------------------------------------
Public Function NewPrintExtensionFromBaseContinue()
On Error GoTo mcrNewPrintExtensionFromBaseContinue_Err


DoCmd.SetWarnings False

    If (Eval("[Forms]![frmNewPrintExtension]![NewPrintExtension] Is Null")) Then
        Beep
        MsgBox "You MUST enter a valid print extension", vbCritical, "PRINT EXTENSION"
        
        Exit Function
    End If
    ' qryPrintExtItemUPDATE
    DoCmd.OpenQuery "qryPrintExtItemUPDATE", acViewNormal, acEdit
    ' qryPrintExtItemCatalogUPDATE
    DoCmd.OpenQuery "qryPrintExtItemCatalogUPDATE", acViewNormal, acEdit
    ' qryPrintExtItemCategoryUPDATE
    DoCmd.OpenQuery "qryPrintExtItemCategoryUPDATE", acViewNormal, acEdit
    ' qryPrintExtItemCategoryUPDATEDesign
    DoCmd.OpenQuery "qryPrintExtItemCategoryUPDATEDesign", acViewNormal, acEdit
    ' qryPrintExtItemCategoryDeleteNull
    DoCmd.OpenQuery "qryPrintExtItemCategoryDeleteNull", acViewNormal, acEdit
    ' qryPrintExtBOMBillOfMtlUPDATE
    DoCmd.OpenQuery "qryPrintExtBOMBillOfMtlUPDATE", acViewNormal, acEdit
    ' qryPrintExtBOMRoutingsInterfaceUPDATE
    DoCmd.OpenQuery "qryPrintExtBOMRoutingsInterfaceUPDATE", acViewNormal, acEdit
    ' qryPrintExtQASpecHeaderUPDATE
    DoCmd.OpenQuery "qryPrintExtQASpecHeaderUPDATE", acViewNormal, acEdit
    ' qryPrintExtQASpecElementUPDATE
    DoCmd.OpenQuery "qryPrintExtQASpecElementUPDATE", acViewNormal, acEdit
    ' clear BOM table
    DoCmd.OpenQuery "qryClearBOMInventoryCompsInterface", acViewNormal, acEdit
    ' clear op res table
    DoCmd.OpenQuery "qryClearBOMOpResourcesInterface", acViewNormal, acEdit
    ' clear op seq table
    DoCmd.OpenQuery "qryClearBOMOpSequencesInterface", acViewNormal, acEdit
    ' clear attachments table
    DoCmd.OpenQuery "qryClearItemInterfaceAttachments", acViewNormal, acEdit
    
    ' Set new print extension on main menu
    Forms!frmMainMenu!NewPrintExtension = Forms!frmNewPrintExtension!NewItemNumber
    
    DoCmd.Close acForm, "frmNewPrintExtension"
    
    DoCmd.OpenForm "frmItemInterface", acNormal, "", "", , acNormal
    
    'DoCmd.Requery ""
    ' Disallow category edits
    Forms!frmItemInterface!frmItemInterfaceCategory.Form.AllowEdits = False
    ' Disallow catalog edits on main form
    Forms!frmItemInterface!frmItemInterfaceCatalog.Form.AllowEdits = False
    'DoCmd.Requery "frmPremiseItemDetailsPrintExt"
    ' Update known values
    'DoCmd.RunMacro "mcrNewItemExtensionUpdate", , ""
    Call NewItemExtensionUpdate
    ' Open catalog elements that can be edited
    'DoCmd.OpenForm "frmPrintExtItemInterfaceCatalog", acNormal, "", "", , acNormal
    ' Turn on "Edit" catalog button
    'Forms!frmItemInterface!frmItemInterfaceCatalog.Form!cmdEdit.Visible = True
    
        ' frmItemInterfaceCatalog - Yes
        Forms!frmItemInterface!frmItemInterfaceCatalog.Form.Visible = True

        ' frmItemInterfaceCategory - Yes
        Forms!frmItemInterface!frmItemInterfaceCategory.Form.Visible = True

        ' frmItemInterfaceBOM - No
        Forms!frmItemInterface!frmItemInterfaceBOM.Form.Visible = False

        ' frmItemInterfaceRouting - No
        Forms!frmItemInterface!frmItemInterfaceRouting.Form.Visible = False


    
DoCmd.SetWarnings True


mcrNewPrintExtensionFromBaseContinue_Exit:
    Exit Function

mcrNewPrintExtensionFromBaseContinue_Err:
    MsgBox Error$
    Resume mcrNewPrintExtensionFromBaseContinue_Exit

End Function
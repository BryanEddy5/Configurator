Option Compare Database
Option Explicit
'------------------------------------------------------------
' mcrCopyItem
'
'------------------------------------------------------------
Function CopyItem() As Boolean
On Error GoTo mcrCopyItem_Err
    DoCmd.SetWarnings False

    ' Clear all item tables
    DoCmd.RunMacro "mcrClearItem", , ""
    ' COPY SELECTED ITEM
    ' qryCopyBOMBillOfMtlsInterface
    DoCmd.OpenQuery "qryCopyBOMBillOfMtlsInterface", acViewNormal, acEdit
    ' qryCopyBOMInventoryCompsInterface
    DoCmd.OpenQuery "qryCopyBOMInventoryCompsInterface", acViewNormal, acEdit
    ' qryCopyBOMOpSequencesInterface
    DoCmd.OpenQuery "qryCopyBOMOpSequencesInterface", acViewNormal, acEdit
    ' qryCopyBOMRoutingsInterface
    DoCmd.OpenQuery "qryCopyBOMRoutingsInterface", acViewNormal, acEdit
    ' qryCopyItemInterface
    DoCmd.OpenQuery "qryCopyItemInterface", acViewNormal, acEdit
    ' qryCopyItemInterfaceAttachments
    DoCmd.OpenQuery "qryCopyItemInterfaceAttachments", acViewNormal, acEdit
    ' qryCopyItemInterfaceCatalog
    DoCmd.OpenQuery "qryCopyItemInterfaceCatalog", acViewNormal, acEdit
    ' qryCopyItemInterfaceCategory
    DoCmd.OpenQuery "qryCopyItemInterfaceCategory", acViewNormal, acEdit
    ' qryCopyQASpecHeaderInterface
    DoCmd.OpenQuery "qryCopyQASpecHeaderInterface", acViewNormal, acEdit
    ' qryCopyQASpecElementInterface
    DoCmd.OpenQuery "qryCopyQASpecElementInterface", acViewNormal, acEdit
    ' Update catalog group name
    DoCmd.OpenQuery "qryCatalogGroupNameUpdate", acViewNormal, acEdit
    DoCmd.OpenForm "frmItemInterface", acNormal, "", "", , acNormal
    DoCmd.Close acForm, "frmNewItemNumber"
    
    CopyItem = True

    DoCmd.SetWarnings True
mcrCopyItem_Exit:
    Exit Function

mcrCopyItem_Err:
    MsgBox Error$
    Resume mcrCopyItem_Exit

End Function
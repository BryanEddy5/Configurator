Option Explicit
Option Compare Database

'------------------------------------------------------------
' mcrClearItem
'
'------------------------------------------------------------
Function mcrClearItem()
On Error GoTo mcrClearItem_Err

    ' qryClearItemInterface
    DoCmd.OpenQuery "qryClearItemInterface", acViewNormal, acEdit
    ' qryClearItemInterfaceCatalog
    DoCmd.OpenQuery "qryClearItemInterfaceCatalog", acViewNormal, acEdit
    ' qryClearItemInterfaceCategory
    DoCmd.OpenQuery "qryClearItemInterfaceCategory", acViewNormal, acEdit
    ' qryClearBOMBillOfMtlsInterface
    DoCmd.OpenQuery "qryClearBOMBillOfMtlsInterface", acViewNormal, acEdit
    ' qryClearBOMInventoryCompsInterface
    DoCmd.OpenQuery "qryClearBOMInventoryCompsInterface", acViewNormal, acEdit
    ' qryClearBOMOperationSetupParameters
    DoCmd.OpenQuery "qryClearBOMOperationSetupParameters", acViewNormal, acEdit
    ' qryClearBOMOpResourcesInterface
    DoCmd.OpenQuery "qryClearBOMOpResourcesInterface", acViewNormal, acEdit
    ' qryClearBOMOpSequencesTempInterface
    DoCmd.OpenQuery "qryClearBOMOpSequencesTempInterface", acViewNormal, acEdit
    ' qryClearBOMOpSequencesInterface
    DoCmd.OpenQuery "qryClearBOMOpSequencesInterface", acViewNormal, acEdit
    ' qryClearBOMRoutingsInterface
    DoCmd.OpenQuery "qryClearItemInterfaceAttachments", acViewNormal, acEdit
    ' qryClearItemInterfaceAttachments
    DoCmd.OpenQuery "qryClearBOMRoutingsInterface", acViewNormal, acEdit
    ' qryClearItemInterfaceQASpecs
    DoCmd.OpenQuery "qryClearItemInterfaceQASpecs", acViewNormal, acEdit
    ' qryClearItemInterfaceQASpecElements
    DoCmd.OpenQuery "qryClearItemInterfaceQASpecElements", acViewNormal, acEdit
    ' qryClearItemCopied
    DoCmd.OpenQuery "qryClearOracleItemInterfaceCopied", acViewNormal, acEdit


mcrClearItem_Exit:
    Exit Function

mcrClearItem_Err:
    MsgBox Error$
    Resume mcrClearItem_Exit

End Function
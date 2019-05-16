Option Compare Database
Option Explicit
'------------------------------------------------------------
' mcrNewItemExtensionUpdate
'
'------------------------------------------------------------
Public Function NewItemExtensionUpdate()
On Error GoTo mcrNewItemExtensionUpdate_Err

Forms!frmItemInterface!frmPremiseItemDetailsPrintExt.Form.Requery

DoCmd.RunSQL "DELETE FROM tblQAItemSpecElements WHERE ElementName = 'LABEL TYPE'"


    ' Set FiberType1 on main menu
    Forms!frmMainMenu!FiberType1 = Forms!frmItemInterface!frmPremiseItemDetails.Form!FiberType1
    ' Update design code
    DoCmd.OpenQuery "qryNewItemExtUpdateCatalogDesign", acViewNormal, acEdit
    ' Update customer version
    DoCmd.OpenQuery "qryNewItemExtUpdateCatalogCustomer", acViewNormal, acEdit
    ' Update customer part number
    DoCmd.OpenQuery "qryNewItemExtUpdateCatalogCustPart", acViewNormal, acEdit
    ' Append Listing Spec
    'DoCmd.OpenQuery "qryNewItemListing", acViewNormal, acEdit
    ' Append E Number Spec
    'DoCmd.OpenQuery "qryNewItemExtEnumber", acViewNormal, acEdit
    If (IsNull(Forms!frmItemInterface!frmPremiseItemDetailsPrintExt.Form!CustomerRev) = False And Forms!frmItemInterface!frmPremiseItemDetailsPrintExt.Form!CustomerRev <> "") Then
        ' Update Customer Revision
        DoCmd.OpenQuery "qryNewItemExtUpdateCatalogCustPartRev", acViewNormal, acEdit
    End If
    DoCmd.OpenQuery "QryNewItemNumberUpdateCatalog_TightTolerance_Fiber"
    
    DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryFiberVendor"
    ' Append "LABEL LISTING"
    'DoCmd.OpenQuery "qryNewItemExtListingLabel", acViewNormal, acEdit
    'DoCmd.Requery ""
    'DoCmd.RefreshRecord
    Forms!frmItemInterface.Refresh
    
    Call AddSpecs
    
mcrNewItemExtensionUpdate_Exit:
    Exit Function

mcrNewItemExtensionUpdate_Err:
    MsgBox Error$
    Resume mcrNewItemExtensionUpdate_Exit

End Function
Option Explicit
Option Compare Database

Function AddSpecs()
Dim ENumber$
'
If [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![FiberType1] Like "fbr*" Then

    DoCmd.OpenQuery "qryNewItemExtUpdateFiberSpecs"

End If

If [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![FiberType2] Like "fbr*" Then

    DoCmd.OpenQuery "qryNewItemExtUpdateFiber2Specs"

End If

If [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![FiberType3] Like "fbr*" Then

    DoCmd.OpenQuery "qryNewItemExtUpdateFiber3Specs"

End If

If Mid([Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]!DesignCode, 8, 1) = "D" And Left([Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]!DesignCode, 1) = "w" Then
    DoCmd.OpenQuery "qryNewItemExtUpdateFiberSpecsStackedRibbon_1300"
    DoCmd.OpenQuery "qryNewItemExtUpdateFiberSpecsStackedRibbon_850"

End If



'Adds copper conductive specs
DoCmd.OpenQuery "qryNewItemExtUpdateCopperUnitSpecs"
DoCmd.OpenQuery "qryNewItemExtUpdateCopperUnitSpecs_ResistanceTest"

DoCmd.RunSQL "DELETE FROM tblQAItemSpecElements WHERE ElementName = 'LABEL TYPE' or ElementName = 'E NUMBER' or ElementName = 'CABLE LISTING'"
'Add label specs


    DoCmd.OpenQuery "qryQASpecENumber", acViewNormal, acEdit
    DoCmd.OpenQuery "qryQASpecCableListing", acViewNormal, acEdit
    DoCmd.OpenQuery "qryNewItemListing", acViewNormal, acEdit

    
    'Delete any label specs that have an empty value
   DoCmd.RunSQL "DELETE FROM tblQAItemSpecElements WHERE (ElementName = 'CABLE LISTING' OR ElementName = 'Print Line 1' OR ElementName = 'Label Type' OR ElementName = 'E Number') AND TargetValue IS NULL OR TargetValue = '' OR TargetValue ='ERROR' "

End Function
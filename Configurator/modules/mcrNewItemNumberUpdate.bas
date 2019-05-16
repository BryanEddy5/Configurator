Option Compare Database
Option Explicit
Dim CableProperties As New clsCableQualities
Dim CableConstruction As New clsCableConstruction

Public Function getCatalogOD(ItemNo As String) As String
Dim OdOne$, OdTwo$

OdOne = Mid(ItemNo, 7, 1)
OdTwo = Mid(ItemNo, 8, 1)


If IsNumeric(OdOne) And IsNumeric(OdTwo) Then

    getCatalogOD = OdOne & "." & OdTwo & "MM"

End If


End Function

Public Function getItemNumberTightbuffer(OracleItemNumber As String) As String
Dim pos%, length%, ItemNumber$, armorSuffix$

armorSuffix = "IA"

If InStr(OracleItemNumber, armorSuffix) > 0 Then
    pos = InStr(OracleItemNumber, armorSuffix)
    length = Len(armorSuffix) - 1
    ItemNumber = Left(OracleItemNumber, pos - length)
Else
    ItemNumber = OracleItemNumber
End If

getItemNumberTightbuffer = ItemNumber



End Function

'------------------------------------------------------------
' mcrNewItemNumberUpdate
'
'------------------------------------------------------------
Public Function NewItemNumberUpdate()
On Error GoTo mcrNewItemNumberUpdate_Errs

DoCmd.SetWarnings False



    ' Set FiberType1 on main menu
    Forms!frmMainMenu!FiberType1 = Forms!frmItemInterface!frmPremiseItemDetails.Form!FiberType1
    ' CATALOG UPDATES
    ' Update design code
    DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogDesign", acViewNormal, acEdit
    ' Update customer version
    DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogCustomer", acViewNormal, acEdit
    ' Update customer part number
        'If (IsNull(Forms!frmItemInterface!frmPremiseItemDetails.Form!CustomerRev) = False And Forms!frmItemInterface!frmPremiseItemDetails.Form!CustomerRev <> "") Then
        ' Update Customer Revision
        DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogCustPart", acViewNormal, acEdit
        'End If
    ' Update color
    DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogColor", acViewNormal, acEdit
    ' Update Fiber Count
    DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogFiberCount", acViewNormal, acEdit
    If (Left(Forms!frmItemInterface!frmPremiseItemDetails.Form!DesignCode, 1) <> "L" And Left(Forms!frmItemInterface!frmPremiseItemDetails.Form!DesignCode, 1) <> "w") Then
        ' Update Diameter
        DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogOD", acViewNormal, acEdit
    End If
    
    ' Update Cable Type
    DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogCableType", acViewNormal, acEdit
    If (Forms!frmItemInterface!frmPremiseItemDetails.Form!OracleItemNumber Like "*ia*") Then
        ' Update Cable Type for Armored items
        DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogCableTypeArmor", acViewNormal, acEdit
    End If
    
    ' Update Cable Rating
    DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogCableRating", acViewNormal, acEdit
    ' Update Jacket Color
    DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogJacketColor", acViewNormal, acEdit
'    If (Mid(Forms!frmItemInterface!frmPremiseItemDetails.Form!OracleItemNumber, 2, 1) <> "w") Then
'        ' Update TB Type
        DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogTBType", acViewNormal, acEdit
'    End If
'
    If (Eval("[Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![DesignCode] Not Like ""*/*"" And Mid([Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![DesignCode],6,1)<>""x""")) Then
        ' Update Fiber Type
        DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogFiberType", acViewNormal, acEdit
    End If
    
    If (Forms!frmItemInterface!frmPremiseItemDetails.Form!DesignCode Like "*/*") Then
        ' Update Fiber Type
        DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogFiberTypeMIXED", acViewNormal, acEdit
    End If
    
    If (Eval("Mid([Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![DesignCode],6,1)=""x"" And [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![DesignCode] Not Like ""*/*""")) Then
        ' Update Fiber
        DoCmd.OpenQuery "qryNewItemNumberUpdateCatalogFiberTypeX", acViewNormal, acEdit
    End If
    
    
    ' BOM UPDATES
    ' Jacket Compound
    DoCmd.OpenQuery "qryNewItemNumberUpdateBOMJacket", acViewNormal, acEdit
    ' Color Chip Compound
    DoCmd.OpenQuery "qryNewItemNumberUpdateBOMColorChip", acViewNormal, acEdit
    DoCmd.RunSQL "DELETE * FROM tblBOMInventoryCompsInterface WHERE ComponentItemNumber = 'NONE'"
    
    If (Eval("[Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![DesignCode] Not Like ""*@*""")) Then
        ' Update Bare Fiber [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![Fiber1] Like "fbr*"
        DoCmd.OpenQuery "qryNewItemNumberUpdateBOMFiber", acViewNormal, acEdit
    End If
    
    If (Forms!frmItemInterface!frmPremiseItemDetails.Form!OracleItemNumber Like "*ia*") Then
        ' Update Armor Core
        DoCmd.OpenQuery "qryNewItemNumberUpdateArmorCore", acViewNormal, acEdit
    End If
    
    If Forms!frmItemInterface!frmPremiseItemDetails.Form!OracleItemNumber Like "*bx*" Then
        DoCmd.OpenQuery "qryNewItemNumberUpdateBoxCable"
    End If
    
    If (False) Then
        ' Update 2 Pass Cable's Subunit
        DoCmd.OpenQuery "qryNewItemNumberUpdateBOMSub", acViewNormal, acEdit
    End If
    ' Routing Updates
    ' Routing 1
    DoCmd.OpenQuery "qryNewItemNumberUpdateRouting1", acViewNormal, acEdit
    ' Routing 2
    DoCmd.OpenQuery "qryNewItemNumberUpdateRouting2", acViewNormal, acEdit
    ' K routing step
    DoCmd.OpenQuery "qryNewItemNumberUpdateRoutingQC", acViewNormal, acEdit
    ' Q routing step
    DoCmd.OpenQuery "qryNewItemNumberUpdateRoutingK", acViewNormal, acEdit
    
    If (Left(Forms!frmItemInterface!frmPremiseItemDetails.Form!Routing1, 1) = "J") Then
        ' Update J std op for Routing 1
        DoCmd.OpenQuery "qryNewItemNumberUpdateRoutingJSeries1", acViewNormal, acEdit
    End If
    
    If (Left(Forms!frmItemInterface!frmPremiseItemDetails.Form!Routing2, 1) = "J") Then
        ' Update J std op for Routing 2
        DoCmd.OpenQuery "qryNewItemNumberUpdateRoutingJSeries2", acViewNormal, acEdit
    End If
    
    If (IsNull(Forms!frmItemInterface!frmPremiseItemDetails.Form!Routing2) = True) Then
        ' Update R routing step
        DoCmd.OpenQuery "qryNewItemNumberUpdateRoutingRNumber", acViewNormal, acEdit
    End If
    
    ' Category Update
    ' Update Fiber Count 1
    DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryFiberCount", acViewNormal, acEdit
    If (Right(Forms!frmMainMenu!NewItemNumber, 3) <> "-00") Then
        ' Update Design to "PRINT"
        DoCmd.OpenQuery "qryNewItemNumberUpdateRouting1", acViewNormal, acEdit
    End If
    
    If (Right(Forms!frmMainMenu!NewItemNumber, 3) = "-00") Then
        ' Update Design to "NO PRINT"
        DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryNOPRINT", acViewNormal, acEdit
    Else
        DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryPrint", acViewNormal, acEdit
    End If
    
    If (Eval("[Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![DesignCode] Not Like ""*@*""")) Then
        ' Update Fiber Family 1
        DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryFiberFamily1", acViewNormal, acEdit
        ' Update Product Class
        
    End If
    
    DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryProductClass", acViewNormal, acEdit
    

    
    If (Forms!frmItemInterface!frmPremiseItemDetails.Form!Location = "2") Then
        SubunitUpdate
    End If
    
     DoCmd.RunSQL "UPDATE tblItemCatalogInterface SET ElementValue = NULL WHERE ElementName = 'QTS Print Vendor Data'"
    
    DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryNPT"
    
    DoCmd.OpenQuery "QryNewItemNumberUpdateCatalog_TightTolerance_Fiber"
    
    DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryFiberVendor"
    
    DoCmd.OpenQuery "qryNewItemNumberUpdateCategoryCableLevel"
    
    Call UnprintedUpdate
    Call BomUsage
    Call RibbonDescUpdate
    Call Update_BOM_Single_Pass_Cables
    
            'Creates the -OSP Oracle description for Ribbon items
    If Not Forms!frmItemInterface!frmPremiseItemDetails.Form!OracleItemNumber Like "rbn*osp" Then
        Call AddSpecs
    End If
    
DoCmd.SetWarnings True


mcrNewItemNumberUpdate_Exit:
    Exit Function

mcrNewItemNumberUpdate_Errs:
    MsgBox Error$
    Resume Next 'mcrNewItemNumberUpdate_Exit

End Function

Function RibbonDescUpdate()
Dim strSQL As String
Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection
Dim FiberDesc As String
Dim strDescription As String



Set cnn = CurrentProject.connection


rst.Open ("SELECT ItemNumber FROM tblItemInterface WHERE OrganizationCode = 'CAB'"), cnn, adOpenStatic, adLockReadOnly

Do Until rst.EOF


    CableProperties.Load (rst!ItemNumber)
    
    CableConstruction.Load (CableProperties.Base)
    
    
    
    If CableConstruction.TBTypeID = 8 Then
    
        Call UpdateRibbonCatalogPeelable(rst!ItemNumber, CableProperties.RibbonHighCure, CableProperties.TBLetterIndicator)
    
    End If


    
    If rst!ItemNumber Like "rbn*osp" Then
        
       
        'Creates the -OSP Oracle description for Ribbon items
        rst.Close
        
        
        
        Set cnn = CurrentProject.connection
        
        'Find Fiber information for description string
        rst.Open "SELECT FiberValue FROM tblFiberElementValue WHERE FiberLetter = '" & Mid(CableProperties.ItemNo, 6, 1) & "'", cnn, adOpenKeyset, adLockOptimistic
        FiberDesc = rst!FiberValue
        rst.Close
        
        'Identify if the item requires high cure
        
        strSQL = "UPDATE tblItemInterface SET Description = '"
        strDescription = "OUTSIDE PROCESSING "
        
        If CableProperties.TBLetterIndicator = 11 Then
            strDescription = "2 LAYER MATRIX,NON-PEELABLE," & strDescription
        End If
        
        If CableProperties.RibbonHighCure = True Then
            strDescription = "HIGH CURE 85%,PEELABLE," & strDescription
        Else
            
        End If
        strDescription = strDescription & "," & Left(CableProperties.ItemNo, 8)
        strDescription = strDescription & "," & CInt(Mid(CableProperties.ItemNo, 3, 3)) & "F BARE RIBBON"
        strDescription = strDescription & "," & FiberDesc
        strDescription = strDescription & ",FIBER (" & CableProperties.Fiber1 & ") SUPPLIED BY AFL '"
        
        strSQL = strSQL & strDescription
        
        DoCmd.RunSQL strSQL
    
    End If

rst.MoveNext
Loop

End Function


Sub UnprintedUpdate()

DoCmd.RunSQL "DELETE FROM tblQAItemSpecElements WHERE ElementName = 'Cable Listing' OR ElementName = 'Print Line 1' OR ElementName = 'Label Type' OR ElementName = 'E Number'"


End Sub


Function UpdateProductClass(Cat3 As String, Cat5 As String, FiberMode As String, Fiber2 As String, item As String, ItemNo As String, Optional Furcation As String) As String

Dim hold(4) As String
Dim IsWtc As Boolean
Dim wTcHold$()





    hold(0) = "Cable"
    hold(1) = "Premise"
    
    If item Like "wtc*" Or item Like "dnt*" Then
        hold(1) = "WTC"
        IsWtc = True
        wTcHold = Split(ItemNo, "-")
        
        If Mid(wTcHold(6), 3, 1) = "S" Then hold(3) = "Armored"

    End If
    

    
    If IsNull(Cat3) Then
        hold(2) = ""
    Else
        hold(2) = Cat3
    End If
        
    Select Case Cat3
        Case "SUB-UNIT"
            If ItemNo Like "em004?09*" Then
                UpdateProductClass = "Cable.BLKJKT.SUB-UNIT.."
            Else
                UpdateProductClass = "Cable.Premise.SUB-UNIT.."
            End If
            Exit Function
        Case "CM-Filler"
            UpdateProductClass = "Cable.Premise.CM-Filler.."
            Exit Function
    End Select

    
    If ItemNo Like "*@*" Then
        hold(3) = "Other"
        hold(4) = "Furcation"
        UpdateProductClass = Join(hold, ".")
        Exit Function
    End If
    
    If Fiber2 = "none" Or IsNull(Fiber2) Then
        hold(3) = FiberMode
    Else
        hold(3) = "Mixed"
    End If
    
    
    If IsNull(Cat5) Then
        hold(4) = ""
    ElseIf Cat3 = "armored" Then
        hold(4) = FindItemData(Left(item, 10), "Oracle_Product_Class_5")
    Else
        hold(4) = Cat5
    
    End If
    
    'Applied for Loose Tube armored cables that are not Premise designs
    If ItemNo Like "l*" And Not IsWtc And (hold(4) = "" Or IsNull(hold(4))) Then
        UpdateProductClass = "Cable.Premise.Armored." & FiberMode & ".LooseTube"
        Exit Function
    End If

    UpdateProductClass = Join(hold, ".")


End Function

Sub UpdateRibbonCatalogPeelable(item As String, highCure As Boolean, RibbonPeelType As Integer)
Dim Descriptor As String
Dim PeelValue As Integer

    If RibbonPeelType <> 11 Then
        Descriptor = "PEELABLE"
        PeelValue = 2
    Else
        Descriptor = "NON-PEELABLE"
        PeelValue = 1
    End If
    
    'DoCmd.SetWarnings False
    DoCmd.RunSQL "delete from tblQAItemSpecElements where ElementName = 'PEELABLE_RIBBON'"
    
    DoCmd.RunSQL "UPDATE tblItemCatalogInterface SET ElementValue = '" & Descriptor & "' WHERE ItemNumber = '" & item & "' AND ElementName = 'SPEC'"
    DoCmd.RunSQL "UPDATE tblItemCatalogInterface SET ElementValue = '" & Descriptor & "' WHERE ItemNumber = '" & item & "' AND ElementName = 'NOTE 1'"
    
    DoCmd.RunSQL "INSERT INTO tblQAItemSpecElements(SpecName, ElementName,UpperSpecLimit,LowerSpecLimit) VALUES ('" & item & "','PEELABLE_RIBBON'," & PeelValue & "," & PeelValue & ")"




End Sub
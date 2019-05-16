Option Compare Database
Option Explicit
Public Const PI As Double = 3.1459265359
Dim dblColorPercentage As Double
Dim dblHelixFactor As Double
Dim strDesignCode As String
Dim SpecificGravity As Double
Dim Extrusion As New clsExtrusionCompounds
Public Enum BomTable
    TempBOM
    BOM
End Enum





Function TBUpdate()

If Forms!frmItemInterface!frmItemInterfaceBOM.Form!ComponentItemNumber Like "pt*" Or Forms!frmItemInterface!frmItemInterfaceBOM.Form!ComponentItemNumber Like "rbn*" _
Or Forms!frmItemInterface!frmItemInterfaceBOM.Form!ComponentItemNumber Like "SWR*" Then
    DoCmd.OpenForm "frmPremiseTBReplace"
End If



'If CurrentProject.AllForms("frmPremiseTBReplace").IsLoaded Then
'Forms("frmPremiseTBReplace").SetFocus
'End If
End Function


Sub SubunitUpdate()

On Error GoTo ErrorHandler:

    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim intCount As Integer
    Dim strComponent As String
    



If [Forms]![frmItemInterface]![frmPremiseItemDetailsCopied].[Form]![OracleItemNumber] Like "*ia" Then
    Exit Sub
End If
    
    Set cnn = CurrentProject.connection
    intCount = 0

    
        rst.Open "SELECT * FROM tblBOMInventoryCompsInterface ORDER BY ItemSeqNumber", cnn, adOpenKeyset, adLockOptimistic
        rst.Filter = "ComponentItemNumber like 'pe*'"
        rst.MoveFirst
        

    Do While Not rst.EOF
        
        
    
        Call ComponentUpdate(intCount, strComponent)
        
        With rst
            !ComponentItemNumber = strComponent
            .Update
        End With
        
        intCount = intCount + 1
        
        rst.MoveNext
    Loop


    Set rst = Nothing
    Set cnn = Nothing

ErrorExit:
    Exit Sub

ErrorHandler:
    If Err.Number = -2147217887 Then
        MsgBox "Check to ensure the Item's units series has subcomponets", vbCritical, "Missing Subunits"
    ElseIf Err.Number = 3021 Then
        MsgBox "There are more or less subs in the copied item than in the loaded item." & vbCrLf & "Verify BOM subcomponents and finished cables fiber type", vbCritical, "Missing subunits"
    Else
        MsgBox Err.Number & " " & Err.Description
    End If
    Resume ErrorExit:


End Sub


Sub ComponentUpdate(intRecord As Integer, strComponent As String)
'Function creates a table to pull components from

    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim strTable As String
    Dim intSubCount As Integer
    Dim intTotalSubCount As Integer


    

    strTable = "tblBOMSubunitComponents"
    'Set MyDB = CurrentDb

    
    
    'Creates table with BOM on the first pass
    If intRecord = 0 Then
        DoCmd.SetWarnings False
        DoCmd.OpenQuery "qryLOADBOMSubunitComponent"
        'Call TableTurnOffDataSubSheet(strTable)
        DoCmd.SetWarnings True
    End If

    
    Set cnn = CurrentProject.connection
    
    rst.Open "tblBOMSubunitComponents", cnn, adOpenKeyset, adLockOptimistic
    intTotalSubCount = rst.RecordCount
    
    rst.Filter = "UnitNumber <> 0"
    intSubCount = rst.RecordCount
    
        
    If intRecord < intSubCount Then
        
        rst.move intRecord
        
        strComponent = rst![Oracle]
    
    End If
    

    If intRecord >= intSubCount Then
        rst.Filter = adFilterNone
        rst.Filter = "[TB Color Series] = 'filler'"
        'Debug.Print rst.RecordCount
        'rst.Move 'intRecord - 2
        
        strComponent = rst![Oracle]
    
    End If



    Set rst = Nothing
    Set cnn = Nothing
    'Set MyProperty = Nothing
    'Set MyDB = Nothing


End Sub


Sub RibbonUpdate()


    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim db As DAO.Database
    Dim intCount As Integer
    Dim strComponent As String
    



If [Forms]![frmItemInterface]![frmPremiseItemDetailsCopied].[Form]![OracleItemNumber] Like "*ia" Then
    Exit Sub
End If
    
    Set cnn = CurrentProject.connection
    Set db = CurrentDb
    intCount = 0

    
        rst.Open "SELECT * FROM tblBOMInventoryCompsInterface ORDER BY ItemSeqNumber", cnn, adOpenKeyset, adLockOptimistic
        rst.Filter = "ComponentItemNumber like 'rbn*'"
        rst.MoveFirst
        

    Do While Not rst.EOF
        
        
        With rst
            !ComponentItemNumber = strComponent
            .Update
        End With
        
        rst.MoveNext
    Loop


    Set rst = Nothing
    Set cnn = Nothing
    


End Sub


Function BomUsage()

'On Error GoTo ErrorHandler:

    Dim cnn As ADODB.connection
    Dim rstInterfaceBOM As New ADODB.Recordset
    Dim intLocation As Integer
    Dim rstMaterial As New ADODB.Recordset
    Dim ComponentMaterial As String
    Set cnn = CurrentProject.connection
    Dim dblMaterialDensity As Double
    Dim beanSkipLoop As Boolean
    Dim dblComponentQuantity As Double
    Dim strCableFamily As String
    Dim dblVolume As Double
    Dim strItemNumber As String
    Dim beanCablingCheck As Boolean
    Dim Count  As Integer
    Dim strColorChip As String

    
    
    DoCmd.SetWarnings False
    
    Count = 1
    
    
    strDesignCode = Forms!frmItemInterface!frmPremiseItemDetails.Form!DesignCode
    
    strItemNumber = [Forms]![frmMainMenu]![NewItemNumber]
    strCableFamily = Left(Forms!frmItemInterface!frmPremiseItemDetails.Form!DesignCode, 1)
    intLocation = Int(Forms!frmItemInterface!frmPremiseItemDetails.Form!Location)
    
   
    
    If Extrusion.Load(strItemNumber) = 0 Then
        MsgBox "Cable information is missing", vbCritical
        GoTo ErrorHandler
    End If
    
    strColorChip = Extrusion.ColorChip
    
    'Calulate volume of 1 meter of cable
    dblVolume = VolumeCalculator(strItemNumber, strDesignCode, 0)
    Call FurcationTubeCheck(intLocation)
    Call ReSequenceBOM(BOM)
    'Set Compound available to 100%
    'Check for cabling step
    beanCablingCheck = CablingCheck
    
    
    
    Set cnn = CurrentProject.connection
    
    rstInterfaceBOM.Open "SELECT * FROM tblBOMInventoryCompsInterface ORDER BY ComponentItemNumber DESC", cnn, adOpenKeyset, adLockOptimistic
    rstInterfaceBOM.MoveLast
    rstInterfaceBOM.MoveFirst
    

    'Loop through each component to calculate the usage rate
    Do Until rstInterfaceBOM.EOF = True
    
    'Current Component to calculate usage
    ComponentMaterial = rstInterfaceBOM!ComponentItemNumber

        rstMaterial.Open "SELECT SpecificGravity FROM dbo_tblMaterialDensity WHERE Compound = '" & ComponentMaterial & "'", cnn, adOpenKeyset, adLockOptimistic
 
        
        'checks if material exists in the dbo_tblMaterialDensity table
        If rstMaterial.EOF = True And Not ComponentMaterial Like "ccp*" Then
            
            beanSkipLoop = True
        Else
            beanSkipLoop = False
        End If
        
        'Identifies components that aren't in the Specific Grabity table
        If beanSkipLoop = True Then
        
            'Finds Subunits and Tight Buffers with a cabling step to update the usage rate
            If beanCablingCheck = True And (Left(ComponentMaterial, 2) = "PE" Or Left(ComponentMaterial, 2) = "PT") Then
                rstInterfaceBOM!ComponentQuantity = dblHelixFactor
                rstInterfaceBOM.Update
            ElseIf beanCablingCheck = False And (Left(ComponentMaterial, 2) = "PE" Or Left(ComponentMaterial, 2) = "PT") Then
                rstInterfaceBOM!ComponentQuantity = 1
            End If
            
            'Finds CenterMembers and Furcation Tubes
            If Left(ComponentMaterial, 3) = "FRP" Or Left(ComponentMaterial, 2) = "PM" Or Left(ComponentMaterial, 3) = "WBT" Then
                'updates Furcation Tubes and center members quantity to 1
                rstInterfaceBOM!ComponentQuantity = 1
                rstInterfaceBOM.Update
            End If
            
            'Finds Zipcord Furcation CAbles and updates the Quantity and Usage to 2
            If Left(ComponentMaterial, 3) = "FRC" Then
                If strCableFamily = "z" Or strCableFamily = "d" Then
                    rstInterfaceBOM!ComponentQuantity = 2
                    rstInterfaceBOM!CountPerUOM = 2
                    rstInterfaceBOM.Update
                Else
                    rstInterfaceBOM!ComponentQuantity = 1
                    rstInterfaceBOM!CountPerUOM = 1
                    rstInterfaceBOM.Update
                End If
            End If
            
            'Identifies aluminum component in BOM
            If Left(ComponentMaterial, 3) = "tia" Then
                rstInterfaceBOM!ComponentQuantity = ArmorUsageRate(Forms!frmItemInterface!frmPremiseItemDetails.Form!Base) 'Finds armor usage rate
            End If
        rstMaterial.Close
        Else
        
                       
            If Left(ComponentMaterial, 3) <> "ccp" Then
                dblMaterialDensity = rstMaterial!SpecificGravity
            End If

                
            

            
            'Identifies aramid in the BOM and applies formula
            Select Case Left(ComponentMaterial, 3)
                Case "ARA", "RIP", "BIN", "MYL", "FGL"
                    'Assigns value of 1 for items without UOM, 1 is default setting in Oracle
                    If IsNull(rstInterfaceBOM!CountPerUOM) = True Or rstInterfaceBOM!CountPerUOM = "" Then
                        rstInterfaceBOM!CountPerUOM = 1
                        rstInterfaceBOM.Update
                    End If
                    'Open Lay Aramid, Binders Ripcords
                    dblComponentQuantity = CDbl(rstInterfaceBOM!CountPerUOM) * dblMaterialDensity
                    dblComponentQuantity = HelixForCabling(ComponentMaterial, dblComponentQuantity, intLocation)
                    rstInterfaceBOM!ComponentQuantity = CheckUsageVariance(rstInterfaceBOM!ComponentQuantity, dblComponentQuantity, ComponentMaterial)
                    rstInterfaceBOM.Update

            End Select
            
        
            
            If Extrusion.CompoundCheck(ComponentMaterial) Then

            'ID's compounds in BOM

                If ComponentMaterial = Extrusion.ColorChip Then
                    dblComponentQuantity = dblVolume * Extrusion.Color_Chip_Concentration * Extrusion.SpecificGravity / 1000
                ElseIf ComponentMaterial = Extrusion.Additive Then
                    dblComponentQuantity = dblVolume * Extrusion.AdditivePercentage * Extrusion.SpecificGravity / 1000
                ElseIf ComponentMaterial = Extrusion.Additive2 Then
                    dblComponentQuantity = dblVolume * Extrusion.AdditivePercentage2 * Extrusion.SpecificGravity / 1000
                ElseIf ComponentMaterial = Extrusion.JacketMaterial Then
                   'Calculates the usage ratee (Volume * Compound% * MaterialDensity /1000 to convert from specific gravity to density
                    dblComponentQuantity = dblVolume * Extrusion.JacketPercentage * Extrusion.SpecificGravity / 1000
                End If
                'Percent of compound leftover for next iteration

                
                rstInterfaceBOM!ComponentQuantity = CheckUsageVariance(rstInterfaceBOM!ComponentQuantity, dblComponentQuantity, ComponentMaterial)
                rstInterfaceBOM.Update
                
            End If
            
            rstMaterial.Close
        End If
    rstInterfaceBOM.MoveNext
    Loop
    'rstMaterial.Close
    rstInterfaceBOM.Close
    cnn.Close
    Forms!frmItemInterface!frmItemInterfaceBOM.Form.Requery
    
    DoCmd.SetWarnings True

ErrorExit:
    Exit Function

ErrorHandler:
    If Err.Number = 94 Then
        MsgBox "Wall thickness, OD, or other construction value is empty", vbCritical
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical
    End If
    Resume ErrorExit:

End Function

Function VolumeCalculator(strItemNumber As String, strItemNo As String, TruncatedCode As Boolean) As Double
    'Function grabs information needed for usage and calculates the volume of the cable's wall

    Dim rst As New ADODB.Recordset
    Dim strCableFamily As String
    Dim dblHeight As Double
    Dim dblOD As Double
    Dim dblWall As Double
    Dim cnn As ADODB.connection
    Dim dblVolume As Double
    Dim strDoubleStackRibbon As String
    
    
    
    
    Set cnn = CurrentProject.connection
    
    rst.Open "SELECT * FROM qryCableConstruction WHERE Oracle='" & strItemNumber & "'", cnn, adOpenKeyset, adLockOptimistic
    
    
    If IsNull(rst!NominalWall) Then
        MsgBox "Material volume cannot be calculated for " & strItemNumber & " due to missing wall thickness.", vbCritical
        Exit Function
    End If
    
    strCableFamily = Left(strItemNo, 1)
    dblOD = CDbl(rst!NominalOD)
    dblWall = CDbl(rst!NominalWall)
    strDoubleStackRibbon = Mid(strItemNo, 7, 1)
        
        
    If TruncatedCode = False Then
        'Identify if there is an existing Helix Factor and store it as a gloabl varabile
        If IsNull(rst!HelixFactor) = True Or rst!HelixFactor = "" Then
            dblHelixFactor = 1
        Else
            dblHelixFactor = rst!HelixFactor
        End If
        
        'Check to see if a colorchip is in BOM.  If not then set colorchip percentage to 0
        If CheckForColorChip = True Then
            dblColorPercentage = rst!Color_Chip_Concentration
        Else
            dblColorPercentage = 0
        End If
        

        
        If IsNull(rst!Height) = True Then
            dblHeight = 0
        Else
            dblHeight = CDbl(rst!Height)
        End If
        
    
        'ID's Oval Cables Cables
        If dblHeight <> 0 And strCableFamily <> "z" Then
            'OD = 1.1284 * sqr(NominalOD*Height) - OD for oblong shaped cables
            dblOD = 1.1284 * Sqr(dblOD * dblHeight)
        
        
            
        End If
        
    End If
    
    'PI()*((OD/2)^2-(ID/2)^2),  ID = OD-(2*NomWall) This formula gives the volume of the cable
    dblVolume = (PI * ((dblOD / 2) ^ 2 - ((dblOD - (2 * dblWall)) / 2) ^ 2))
    
    'Identifies zipcord construction
    If strCableFamily = "z" Then
        dblVolume = 2 * dblVolume
    ElseIf strCableFamily = "w" Then
        dblVolume = Int(strDoubleStackRibbon) * dblVolume
    End If
    
    VolumeCalculator = dblVolume
    
    rst.Close
    cnn.Close
    

End Function

Function CheckUsageVariance(dblCurrentQuantity As Double, dblNewQuantity As Double, ComponentMaterial As String) As Double

    'Check if the calculated usage rate if +/- 15% of what is in the BOM currently
    If dblCurrentQuantity <= dblNewQuantity * 1.15 And dblCurrentQuantity >= dblNewQuantity * 0.85 Then
            CheckUsageVariance = dblNewQuantity
    Else
        If MsgBox("The varaince between the calculated rate and the BOM rate is +/-15%. " & _
            vbCrLf & "Do you want to update the Compound usage rate from " & Round(dblCurrentQuantity, 6) & " to " & Round(dblNewQuantity, 6) & " of " & ComponentMaterial & "?", vbYesNo) = vbYes Then
            
            CheckUsageVariance = dblNewQuantity
        Else
            CheckUsageVariance = dblCurrentQuantity
        End If
    End If

End Function


Function CheckForColorChip() As Boolean


    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    
    If strDesignCode Like "*box*" Then
        CheckForColorChip = False
        Exit Function
    End If
    
    
    
    
    'Deletes the row for colorchip information if there is "NONE" for the colorchip type
    DoCmd.RunSQL "DELETE * FROM tblBOMInventoryCompsInterface WHERE ComponentItemNumber = 'NONE'"
    
    Set cnn = CurrentProject.connection
    
    
    
    rst.Open "SELECT * FROM tblBOMInventoryCompsInterface WHERE left(ComponentItemNumber,3) = 'ccp' ORDER BY ComponentItemNumber DESC", cnn, adOpenKeyset, adLockOptimistic
    
    
    'Check to see if a colorchip is in BOM
    If rst.EOF = True Then
        'No color chip in BOM, but item needs color chip.  Appends color chip and Flags as having a color chip for calculation
        If Extrusion.Color <> "NONE" And IsNull(Extrusion.Color) = False Then
            DoCmd.OpenQuery "qryAPPENDColorChip"
            Call AddColorChip
            CheckForColorChip = True
        'No color chip is needed in BOM
        Else
            CheckForColorChip = False
        End If
    Else
        'Color Chip in BOM and Access Item
        CheckForColorChip = True
    End If
    
    rst.Close
cnn.Close


End Function


Function AddColorChip()
'Finds if current BOM has color chip and inserts or deletes appropriately
    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    Dim intNextItemSeqNumber As Integer
    
    Set cnn = CurrentProject.connection
    
    rst.Open "SELECT * FROM tblBOMInventoryCompsInterface ORDER BY ComponentItemNumber DESC", cnn, adOpenKeyset, adLockOptimistic
    rst.MoveFirst
    intNextItemSeqNumber = rst!ItemSeqNumber + 10
    rst.Close
    
    rst.Open "SELECT * FROM tblBOMInventoryCompsInterface WHERE ComponentItemNumber ='" & Extrusion.ColorChip & "'", cnn, adOpenKeyset, adLockOptimistic
    
    
    rst!ItemSeqNumber = intNextItemSeqNumber
    rst.Update
    rst.Close
    cnn.Close
    
    
    DoCmd.RunSQL "DELETE FROM tblBOMInventoryCompsInterface where ComponentItemNumber = 'NONE'"



End Function

Function HelixForCabling(strComponentMaterial As String, dblComponentQuantity As Double, Location As Integer) As Double

'Function finds if the cable should have a helix factor

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    Dim strStdOp As String
    Dim strSQL As String
    
    'If Location = 1 Then
    '    HelixForCabling = dblComponentQuantity
    '    Exit Function
    'End If
    
    
    Set cnn = CurrentProject.connection
    
    strSQL = " SELECT * FROM tblBOMInventoryCompsInterface INNER JOIN tblBOMOpSequencesInterface "
    strSQL = strSQL & " ON tblBOMInventoryCompsInterface.OperationSeqNumber = tblBOMOpSequencesInterface.OperationSeqNum "
    strSQL = strSQL & "WHERE ComponentItemNumber = '" & strComponentMaterial & "'"
    
    
    rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    
    strStdOp = rst!OperationCode
    
    If Left(strStdOp, 1) = "c" Then
        HelixForCabling = dblHelixFactor * dblComponentQuantity
    Else
        HelixForCabling = dblComponentQuantity
    End If
    
    
    
    rst.Close
    cnn.Close

End Function

Function CablingCheck() As Boolean

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    
    Set cnn = CurrentProject.connection
    
    rst.Open "SELECT OperationCode FROM tblBOMOpSequencesInterface WHERE left(OperationCode,1) = 'c'", cnn, adOpenForwardOnly, adLockReadOnly
    
    If rst.EOF = True Then
        CablingCheck = False
    Else
        CablingCheck = True
    End If
    
    rst.Close
    cnn.Close
    
End Function

Function FurcationTubeCheck(Location As Integer)
    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    
    Set cnn = CurrentProject.connection
    
    rst.Open "SELECT * FROM qryFurcationCheck WHERE Oracle = '" & [Forms]![frmMainMenu]![NewItemNumber] & "'", cnn, adOpenForwardOnly, adLockReadOnly
    
    If rst.EOF = True Or Location = 2 Then
        Exit Function
    End If
    
    'Clear all Tight Buffers and Furcation from BOM
    DoCmd.OpenQuery "qryBOMFurcationCLEAR"
    
    'Append Furcation tube to bom
    DoCmd.OpenQuery "qryAPPENDFurcationTube"
    

    
    rst.Close
    cnn.Close



End Function

Function ArmorUsageRate(Base) As Double

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    Dim ArmorType() As String
    ArmorType = Split(Base, "-")
    
    Set cnn = CurrentProject.connection
    
    rst.Open "SELECT * FROM tblDesignArmor WHERE armor = '" & ArmorType(0) & "'", cnn, adOpenForwardOnly, adLockReadOnly
    
    ArmorUsageRate = rst!armorusage
    'Debug.Print rst!armorusage



End Function

Function ReSequenceBOM(TableType As BomTable)

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    Dim seq As Integer, Table$, qry$
    
    If TableType = TempBOM Then
        Table = "tblItemLoadTempBOM"
    Else
        Table = "tblBOMInventoryCompsInterface"
        
    End If
    
    seq = 10
    
    Set cnn = CurrentProject.connection
    
    qry = "SELECT ItemSeqNumber FROM " & Table & " ORDER BY OperationSeqNumber ASC, ItemSeqNumber ASC"
    
    rst.Open qry, cnn, adOpenForwardOnly, adLockOptimistic
    
    With rst
        .MoveFirst
        Do Until .EOF
            !ItemSeqNumber = seq
            .MoveNext
            seq = seq + 10
        Loop
        .UpdateBatch
    End With


End Function


Function BOM_WeightCalculator(item As String) As Double



Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim strStoredProcedure As String
Dim rst As New ADODB.Recordset
Dim OracleItem As String
Dim bArmor As Boolean
Dim dbWeight As Double
Dim JacketMaterial As String
Dim Base As String
Dim ItemNo As String
Dim ItemData As Scripting.Dictionary
Dim CoreItem As String
'Set ItemDAta = New Scripting.Dictionary


On Error GoTo ErrorHandler:

If Mid(item, 11, 2) = "IA" Then
    CoreItem = FindCoreCable(item)
    Set ItemData = RetrieveItemData(item)
    Base = ItemData.item("Base")
    JacketMaterial = ItemData.item("JacketMaterial")
    ItemNo = ItemData.item("ItemNo")


    bArmor = True
    'OracleItem = FindExistingBOMItem(CoreItem)
    OracleItem = CoreItem
Else
    'OracleItem = FindExistingBOMItem(Item)
    OracleItem = item
End If

'Function needs refining for Armored BOMs



strStoredProcedure = "SELECT Weight FROM usf_BOM_Weight(N'" & OracleItem & "')"

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr



rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic
 
'Debug.Print LTrim(rst!Weight)

If bArmor = True Then
    dbWeight = VolumeCalculator(item, ItemNo, 1) * CompoundSpecificGravity(JacketMaterial)
    dbWeight = dbWeight + ArmorUsageRate(Base) * 1000
    dbWeight = dbWeight + rst!Weight
    dbWeight = Round(dbWeight)
Else
    dbWeight = rst!Weight

End If

'Debug.Print rst!Weight
BOM_WeightCalculator = dbWeight

rst.Close
cnn.Close

ErrorExit:
Exit Function

ErrorHandler:
MsgBox "Error calculating weight.  Please contact product engineering.", vbCritical
Resume ErrorExit


End Function

Function FindExistingBOMItem(item As String) As String

'If the item is a bronze braided or stainless steel braided, then do not choose another item for weight
If item Like "*-??BB" Or item Like "*-??SS" Or item Like "DN*" Then
    FindExistingBOMItem = item
    Exit Function
End If

Dim rst As New ADODB.Recordset
Dim cnn As New ADODB.connection
Dim strSQL As String
Dim provstr$

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr



strSQL = "SELECT * FROM dbo.GetExistingOracleItemByBase('" & item & "')"

rst.Open (strSQL), cnn, adOpenKeyset, adLockOptimistic


FindExistingBOMItem = rst!ExistingItem

rst.Close
cnn.Close




End Function


Function CompoundSpecificGravity(Compound As String) As Double

Dim rstMaterial As New ADODB.Recordset
Dim cnn As ADODB.connection
Set cnn = CurrentProject.connection



rstMaterial.Open "SELECT SpecificGravity FROM dbo_tblMaterialDensity WHERE Compound = '" & Compound & "'", cnn, adOpenForwardOnly, adLockReadOnly

CompoundSpecificGravity = rstMaterial!SpecificGravity

End Function


Function RetrieveItemData(item As String) As Scripting.Dictionary

Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection
Dim strSQL As String
Set cnn = CurrentProject.connection
'Dim Data As Collection
'Set Data = New Collection
Dim Data As Scripting.Dictionary
Set Data = New Scripting.Dictionary

strSQL = "SELECT [Item No], tblCableConstructionReferences.Base, JacketMaterial, DesignCode, NominalOD "
strSQL = strSQL & " FROM (dbo_CableUnion INNER JOIN tblCableConstructionReferences ON dbo_CableUnion.Base = tblCableConstructionReferences.Base) "
strSQL = strSQL & " INNER JOIN tblCableConstructions ON tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID "
strSQL = strSQL & " WHERE dbo_CableUnion.Oracle = '" & item & "'"

rst.Open (strSQL), cnn, adOpenKeyset, adLockOptimistic


Data.Add "Base", rst!Base.value
Data.Add "ItemNo", rst![Item No].value
Data.Add "JacketMaterial", rst!JacketMaterial.value
Data.Add "DesignCode", rst!DesignCode.value
Data.Add "NominalOD", rst!NominalOD.value


Set RetrieveItemData = Data

rst.Close
cnn.Close



End Function

Function FindCoreCable(item As String) As String 'Find core item for Armored cables
Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection
Set cnn = CurrentProject.connection
Dim strSQL As String

strSQL = "SELECT ComponentItemNumber FROM dbo_AFLPRD_BOMInvComp_CAB WHERE ItemSeqNumber = 10 AND AssemblyItemNumber = '" & item & "'"
rst.Open (strSQL), cnn, adOpenKeyset, adLockOptimistic

If rst.EOF Then
    FindCoreCable = Replace(item, "IA", "")
Else
    FindCoreCable = rst!ComponentItemNumber
End If


End Function

Function FindItemFromBase(Base As String) As String

'If the item is a bronze braided or stainless steel braided, then do not choose another item for weight


Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection
Dim strSQL As String
Set cnn = CurrentProject.connection



strSQL = "SELECT dbo_CableUnion_1.Oracle "
strSQL = strSQL & " FROM AFLPRD_INVSysItemCost_CAB INNER JOIN (dbo_CableUnion AS dbo_CableUnion_1 INNER JOIN (dbo_CableUnion INNER JOIN tblCableConstructionReferences ON "
strSQL = strSQL & " dbo_CableUnion.Base = tblCableConstructionReferences.Base) ON dbo_CableUnion_1.Base = tblCableConstructionReferences.Base) ON AFLPRD_INVSysItemCost_CAB.ItemNumber = dbo_CableUnion_1.Oracle "
strSQL = strSQL & " WHERE dbo_CableUnion.Base = '" & Base & "' AND dbo_CableUnion_1.Oracle NOT Like '%osp%'AND dbo_CableUnion_1.Oracle NOT Like '%-__BX%'"
strSQL = strSQL & " ORDER BY dbo_CableUnion_1.Oracle DESC"

rst.Open (strSQL), cnn, adOpenKeyset, adLockOptimistic



FindItemFromBase = rst!Oracle

rst.Close
cnn.Close




End Function




Function Base_BOM_WeightCalculator(Base As String) As Double

'Calculates the weight of Items buy using an SQL procedure to aggregate all materials and summate them



Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim strStoredProcedure As String
Dim rst As New ADODB.Recordset
Dim OracleItem As String
Dim dbWeight As Double
'Set ItemDAta = New Scripting.Dictionary


On Error GoTo ErrorHandler:


'Function to Find an item BOM in oracle
OracleItem = FindItemFromBase(Base)



strStoredProcedure = "DECLARE @return_value int, @Weight decimal(16,4) "
strStoredProcedure = strStoredProcedure & "EXEC    @return_value = [dbo].[CalculateBOMWeight]  N'" & OracleItem & "',@Weight OUTPUT "
strStoredProcedure = strStoredProcedure & "SELECT  'Weight' = @Weight"

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr



rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic
 


dbWeight = rst!Weight


Base_BOM_WeightCalculator = dbWeight

rst.Close
cnn.Close

ErrorExit:
Exit Function

ErrorHandler:
MsgBox "Error calculating weight.  Please contact product engineering.", vbCritical
Resume ErrorExit


End Function

Function Update_BOM_Single_Pass_Cables()
'Function opens BOM table and identifies components to be updated with the appropriate ribbon or tight buffer
''Lacking split fiber logic''

Dim rstTB As New ADODB.Recordset
Dim cnn As ADODB.connection
Dim rstBOM As New ADODB.Recordset
Dim strBOMSQL As String
Dim sNewComponent As String
Dim iLength As Integer
Dim sSearchCriteria As String
Dim TBTypeID As Integer
Dim rst As New ADODB.Recordset
Dim updateSuccessfull As Boolean
Dim CableProperties As New clsCableQualities

updateSuccessfull = False


Set cnn = CurrentProject.connection


rst.Open ("SELECT ItemNumber FROM tblItemInterface WHERE OrganizationCode = 'CAB'"), cnn, adOpenStatic, adLockReadOnly

Do Until rst.EOF

    CableProperties.Load (rst!ItemNumber)
    
    If CableProperties.Location = 1 And CableProperties.ComponentSelection Then
        
    
    
        rstTB.Open "SELECT * FROM qryUpdateBOMTightBuffers", cnn, adOpenStatic, adLockReadOnly
        
        
        If rstTB.RecordCount = 0 Then  'Determines if function should run
            rstTB.Close
            cnn.Close
            MsgBox "There are no available components for BOM substitution.  Create the necessary components.", vbCritical
            Exit Function
        
        ElseIf rstTB.RecordCount > 1 Then
             MsgBox "More than 1 component is available for BOM substitution.  Open BOM and select appropriate component.", vbCritical
             Exit Function
        End If
        
        
        TBTypeID = rstTB!TBTypeID
        
        
        
         If CableProperties.TightBuffered Then
         
        
                iLength = 7                                         'Length of tight buffer parent that needs to be replaced in BOM
                sNewComponent = rstTB!Component                     'Get new component to update BOM with
                sSearchCriteria = "pt%"
                

        Else
        
            If TBTypeID = 9 Then      'SWR cables
            
                iLength = 8                                         'Length of SWRR parent that needs to be replaced in BOM
                sNewComponent = rstTB!Component                     'Get new component to update BOM with
                sSearchCriteria = "SWR%"
            
            ElseIf TBTypeID = 8 Then       'Ribbon Cables  ***NEEDS LOGIC FOR FIBR COUNT****
                iLength = 8                                         'Length of SWRR parent that needs to be replaced in BOM
                sNewComponent = rstTB!Component                     'Get new component to update BOM with
                sSearchCriteria = "RBN%"
            
            ElseIf TBTypeID = 21 Then      'Tight Structure Cables
            
            End If
        End If
        
        
            
        strBOMSQL = "SELECT ComponentItemNumber FROM tblBOMInventoryCompsInterface WHERE ComponentItemNumber LIKE '" & sSearchCriteria _
        & "' AND ComponentItemNumber NOT LIKE 'PT%-%N'" 'PT%-0?N as filler tight buffers that should remain unchanged
        
        
        rstBOM.Open (strBOMSQL), cnn, adOpenForwardOnly, adLockOptimistic
        
        Do Until rstBOM.EOF
            rstBOM!ComponentItemNumber = Left(sNewComponent, iLength) & Mid(rstBOM!ComponentItemNumber, iLength + 1, 7)
            updateSuccessfull = True
            rstBOM.MoveNext
        Loop
        
        If updateSuccessfull = False Then
            MsgBox "BOM component update failed.  Carefully review BOM.", vbCritical
        End If
        
    End If
    rst.MoveNext
Loop

End Function

Function BOM_TightBuffer_Replace(TBTypeID As String)

Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection

Set cnn = CurrentProject.connection

rst.Open ("SELECT BOMFunction FROM tblCableTightBufferReference WHERE TBTypeID = '" & TBTypeID & "'"), cnn, adOpenForwardOnly, adLockReadOnly



End Function
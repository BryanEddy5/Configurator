Option Compare Database
Dim intNumSubFillers As Integer
Dim intNumSub As Integer
Dim intUnitSeriesSubCount As Integer
Dim intNumCopperUnits As Integer
Dim CableConstruction As New clsCableConstruction
Dim SubColorID As String
'Option Explicit
'This procedure creates a new Unit Series for cables with differing subunit colors

Function frmSubunitX()
Dim vstartNumber As Integer
Dim SubCount As Integer
Dim OM As String
Dim strMMPosition As Integer
Dim strSMPosition As Integer
Dim strExtension As String
Dim strSubPositionStandard As Integer




On Error GoTo Error_Handler:

DoCmd.SetWarnings False

'Call Functions to build Unit Series Template
InsertUnitSeries (Forms!frmCreateItem!Base)

CableConstruction.Load (Forms!frmCreateItem!cbSubunit)


strExtension = Mid(Forms!frmCreateItem!ItemNo, 12, 30)
strMMPosition = InStr(strExtension, "MM")
strSMPosition = InStr(strExtension, "SM")


If Forms!frmCreateItem!MixedFiber = 1 Then
    DoCmd.Close acForm, "frmFiberSpec"
Else
End If
    
DoCmd.Close acForm, "frmFiberMode"

Forms!frmCreateItem!SubCount = 1
Forms!frmCreateItem!SubColorCount = 1


'If creating sequential subunits then the first subunit must be blue (1)
If Forms!frmCreateItem!MixedSub = 1 Or Forms!frmCreateItem!MixedTB = 1 Then
    Forms!frmCreateItem!SubColor = 1
End If




'Designates the fiber to use for the first sub
Forms!frmCreateItem!FiberType = Forms!frmCreateItem!Fiber1Type

OM = Forms!frmCreateItem!OM

'Counts the number of subunits inside the cable design
'DoCmd.OpenForm "frmUnitSeriesCount"



'Identifies how many subs will need to be created for each fiber type in mixed fiber cables
If Forms!frmCreateItem!MixedFiber = 1 Then
    Forms!frmCreateItem!Sub1Count = Forms!frmCreateItem!Fiber1Count
    Forms!frmCreateItem!Sub2Count = Forms!frmCreateItem!Fiber2Count
    
    If strMMPosition < strSMPosition Then
        Forms!frmCreateItem!Fiber = Forms!frmCreateItem!Fiber1
        Forms!frmCreateItem!FiberType = Forms!frmCreateItem!Fiber1Type
        OM = Forms!frmCreateItem!OM
        SubCount = Forms!frmCreateItem!Sub1Count / CableConstruction.FiberCount
        strSubPositionStandard = 1
    Else
        Forms!frmCreateItem!Fiber = Forms!frmCreateItem!Fiber2
        Forms!frmCreateItem!FiberType = Forms!frmCreateItem!Fiber2Type
        OM = Forms!frmCreateItem!OM2
        SubCount = Forms!frmCreateItem!Sub2Count / CableConstruction.FiberCount
        strSubPositionStandard = 0
    End If
    
Else
End If


For vstartNumber = 1 To Forms!frmCreateItem!SubNo Step 1

    DoEvents
    
    DoCmd.OpenForm "frmFiberMode"
    
    If Forms!frmCreateItem!MixedFiber = 1 Then
        DoCmd.OpenForm "frmFiberSpec"
    Else
    End If
    
    'Changes color of subunit based on fiber mode (hybrid cables)
    If Forms!frmCreateItem.HybridSubColor = 1 Then
        Select Case OM
            Case "OM1", "OM2"
                Forms!frmCreateItem!SubColor = 2
            Case "OM3", "OM4"
                Forms!frmCreateItem!SubColor = "C"
            Case "SM"
                Forms!frmCreateItem!SubColor = 9
        End Select
        
        If Forms!frmCreateItem!TBY = 1 Then
            Forms!frmCreateItem!TBColor = Forms!frmCreateItem!SubColor
        End If
            
    End If
    
            'DoCmd.OpenForm "frmSubUnitXexisting"             'Determines if the subunit already exists in tables

        

    'If the subunit does not exist then this will create a new sub
        'If Forms!frmSubUnitXexisting!Existing = 0 Then

        If SubUnitExist() = "0" Then 'Finds if subunit exists and records either "0" or the existing subunit

        
            
            'DoCmd.OpenForm "frmSubUnitJacketColor_X"      ' Opens form to update the new sub's jacket colorchip and color
            DoCmd.OpenQuery "qrySubunitUpdate_X"          'Updates Subunit information for creation of new item
           
            If CableConstruction.TightBuffered Then
                Call Forms("frmCreateItem").TBMaterial_Update
            Else
                DoCmd.OpenQuery "qryfrmCreateItemUpdateTBMaterial_NoMaterial"
            End If
            
            
            
            DoCmd.OpenQuery "qrySubunitNewUnitSeriesUpdate_X" 'Updates XX-Unit in TblUnitSeriesCopy
          
          
            DoCmd.OpenQuery "qryNewBasicConstructionAppend" 'Appends to the hold table for display at the end of the cable creation

            

            
    'If Subunit already exists then it will update the information in the table
        Else
        
            DoCmd.OpenQuery "qryUnitSeriesSubunitClear"
            DoCmd.OpenQuery "qryUnitSeriesSubunit_X_Append_New"
            'InsertSubUnit (strSubUnit)
            'Call SubUnitExist
            DoCmd.OpenQuery "qryUnitSeriesCopyItemNoUpdate_X"
            
            
        End If
        
        'Counter to change the criteria to find/create the next loop
  
        Forms!frmCreateItem!SubColorCount = Forms!frmCreateItem!SubColorCount + 1
        Forms!frmCreateItem!SubCount = Forms!frmCreateItem!SubCount + 1
        
        'Leaves the subunit color the same if it is not a mixed color sub
        If Forms!frmCreateItem!HybridSubColor = 1 And Forms!frmCreateItem!TBY = 1 Then
            Forms!frmCreateItem!TBColor = Forms!frmCreateItem!SubColor
        ElseIf Forms!frmCreateItem!MixedSub = 1 Or Forms!frmCreateItem!MixedTB = 1 Or Forms!frmCreateItem!HybridSubColor = 1 Then
            Forms!frmCreateItem!SubColor = Forms!frmCreateItem!SubColorCount
        End If
        
        'If Forms!frmCreateItem!TBY = 1 Then
          '  Forms!frmCreateItem!TBColor = Forms!frmCreateItem!SubColorCount
        'End If
        
        
        'Changes the Fiber to Fiber 2 after the updating for Fiber 1 items has completed
        
        If Forms!frmCreateItem!MixedFiber = 1 Then
            If Forms!frmCreateItem!SubColorCount > SubCount Then
                DoCmd.Close acForm, "frmFiberSpec"
                If strSubPositionStandard = 1 Then
                    Forms!frmCreateItem!Fiber = Forms!frmCreateItem!Fiber2
                    Forms!frmCreateItem!FiberType = Forms!frmCreateItem!Fiber2Type
                    OM = Forms!frmCreateItem!OM2
                Else
                    Forms!frmCreateItem!Fiber = Forms!frmCreateItem!Fiber1
                    Forms!frmCreateItem!FiberType = Forms!frmCreateItem!Fiber1Type
                    OM = Forms!frmCreateItem!OM
                End If
                DoCmd.OpenForm "frmFiberSpec"
            End If
         End If
         
         'Changes iterative steps to appropriate Jacket Color Letter
         
        If Forms!frmCreateItem!MixedTB = 1 Then
            Forms!frmCreateItem!TBColor = Forms!frmCreateItem!SubColor
        Else
            Forms!frmCreateItem!TBColor = 0
        End If
        
         
         SubColorID = Int(Forms!frmCreateItem!SubColor) Mod 12
         
         Select Case SubColorID
            Case 10
                SubColorID = "A"
            Case 11
                SubColorID = "B"
            Case 0
                SubColorID = "C"
        End Select
        
       Forms!frmCreateItem!SubColor = SubColorID
       
       
       'Changes iterative steps to appropriate TB Color Letter
       If Forms!frmCreateItem!TBColor <> 0 And (Forms!frmCreateItem!TBY <> 1 Or IsNull(Forms!frmCreateItem!TBY)) Then
             
             If Forms!frmCreateItem!SubColorCount >= 13 And Forms!frmCreateItem!SubColorCount < 36 Then
                 Forms!frmCreateItem!TBColor = "S"
             Else
                 Forms!frmCreateItem!TBColor = Forms!frmCreateItem!SubColor
             End If
        End If
          
            
 'DoCmd.Close acForm, "frmSubUnitXexisting"
 DoCmd.Close acForm, "frmFiberMode"
 



    Next

DoCmd.Close acForm, "frmFiberSpec"

If Forms!frmCreateItem!MixedFiber = 1 Then
Call BasicCableMixedFiber
End If
DoCmd.SetWarnings True

Error_Exit:
Exit Function

Error_Handler:
rst.Close
cnn.Close
'cnn = Nothing
DoCmd.Close acForm, "frmFiberSpec"
'DoCmd.Close acForm, "frmSubUnitXexisting"
DoCmd.Close acForm, "frmSubUnitNextItemNo"
'DoCmd.Close acForm, "frmSubUnitJacketColor_X"
DoCmd.Close acForm, "frmFiberMode"
'DoCmd.Close acForm, "frmSpecialPrintMatch"
MsgBox Err.Description & " " & Err.Number
Resume Error_Exit


End Function


Sub BasicCableMixedFiber()

Forms!frmCreateItem!Fiber = Forms!frmCreateItem!Fiber2

DoCmd.OpenForm "frmFiberSpec"

DoCmd.OpenQuery "qryUpdateNewBasicCableMixedFiber"
DoCmd.Close acForm, "frmFiberSpec"

End Sub

Public Function SubUnitExist() As String

    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Set cnn = CurrentProject.connection
    
    
    'strItemNo = Forms!frmMainMenu!ItemNo


    strSQL = "SELECT  * FROM (tblCableConstructionReferences INNER JOIN [Product Desrcriptions] ON(tblCableConstructionReferences.SetupID = [Product Desrcriptions].ProductID) AND "
    strSQL = strSQL & "(tblCableConstructionReferences.SetupID = [Product Desrcriptions].ProductID)) INNER JOIN qrySearchAllSubs ON [Product Desrcriptions].[Level 1 Product] = qrySearchAllSubs.Base"
    strSQL = strSQL & " WHERE qrySearchAllSubs.Active <> 0 AND [Fiber Oracle Item] = '" & Forms!frmCreateItem!Fiber & "' AND [New Oracle Part #] not like '%test%'"
    strSQL = strSQL & " AND Mid([qrySearchAllSubs]![Item No],10,1) = '" & Forms!frmCreateItem!SubColor & "'"
    strSQL = strSQL & " AND Mid([qrySearchAllSubs]![Item No],11,2) = '" & [Forms]![frmCreateItem]![TBColor] & [Forms]![frmCreateItem]![TBType] & "'"
    strSQL = strSQL & " AND tblCableConstructionReferences.Base = '" & [Forms]![frmCreateItem]![Base] & "'"
    strSQL = strSQL & " AND (FiberType2 = 'NONE' OR FiberType2 is null) AND SubFiller = 0 AND [Level 1 Product] ='" & Forms!frmCreateItem!cbSubunit & "'"
  

    rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    
    
    If rst.EOF Then
    
    
    SubUnitExist = "0"
    
    Else
    rst.MoveFirst
    
    
    SubUnitExist = rst![New Oracle Part #]
    
    
    
    End If
    

    rst.Close
    Set cnn = Nothing
    
End Function

Public Function InsertUnitSeries(strBase As String)

Dim cnn As ADODB.connection
Dim rst As New ADODB.Recordset
Dim strSQL As String
Set cnn = CurrentProject.connection
Dim rst2 As New ADODB.Recordset
Dim start As Integer

'On Error GoTo ErrorHandler:

DoCmd.RunSQL "DELETE * FROM tblNewPrductConstruction"
DoCmd.OpenQuery "qrySubunitAppend"                      'Appends Base and Jacket to build new item with if necessary

'Create Unit Series to then populate with subunits
CountSubs (strBase)

strSQL = "SELECT UnitID, PrintSpacing, UnitIDType,PrintType,UnitNumber "
strSQL = strSQL & "FROM tblUnitSeriesTemplate INNER JOIN tblCableConstructionReferences "
strSQL = strSQL & "ON tblUnitSeriesTemplate.UnitIDTypeNumber = tblCableConstructionReferences.UnitIDTypeNumber "
strSQL = strSQL & "WHERE UnitNumber <= " & intUnitSeriesSubCount & " AND Base = '" & Forms!frmCreateItem!Base & "' AND UnitNumber <>0  "
rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic


DoCmd.RunSQL "DELETE * FROM tblUnitSeriesCopy"
rst2.Open "tblUnitSeriesCopy", cnn, adOpenKeyset, adLockOptimistic


'Add Unit Series Template to the tblUnitSeriesTemp table
Do Until rst.EOF
    rst2.AddNew
    'If intNumSub = 1 Then
    If intUnitSeriesSubCount = 1 Then
    rst2!UnitNumber = 100
    Else
        rst2![Unit ID] = rst!UnitID
        rst2!UnitNumber = rst!UnitNumber
        rst2![Print Spacing] = rst!PrintSpacing

        rst2![Print type] = rst!PrintType
    End If
    rst2![TB Color Series] = "std"
    rst2!verified = -1
    rst2![Unit Color Series] = Forms!frmCreateItem!UnitSeries
    rst.MoveNext
    rst2.MoveNext
Loop
rst2.UpdateBatch


'Add the number of positions needed to hold fillers and copper units to tblUnitSeriesCopy - for holding
If intNumSubFillers <> 0 Then

    For start = 1 To intNumSubFillers
    
    DoCmd.OpenQuery "qryUnitSeriesFiller"   'Appends the records to tblUnitSeriesCopy to house the filler/copper unit data
    
    
    Next
    
    
    Dim fillerQry$
    fillerQry = "qryUnitSeriesFillerAppend"
    'Checks to see if a filler is available
'    rst3.Open ("SELECT * FROM dbo.vFillersActive"), cnn, adOpenKeyset, adLockOptimistic
'    If rst.EOF Then
'        MsgBox "Filler could not be found.  May need to be created.", vbCritical
'    Else
        DoCmd.OpenQuery fillerQry 'Updates tblUnitSeriesCopy with the fillers
    'End If
End If


rst.Close
rst2.Close
cnn.Close

ErrorExit:
Exit Function

ErrorHandler:
rst.Close
cnn.Close
'MsgBox ("Missing UnitIDTypeNumber in tblCableConstructionReferences")
Resume ErrorExit



End Function

Public Function FillerExist(Base$) As Boolean
'TODO: Check to see if a filler exists and notify the user if it does not

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr

sQry = "SELECT Oracle FROM dbo.CableConstructionReferences INNER JOIN "


End Function



Public Function CountSubs(Base As String) As Integer

'Function to Count the number of subs in a

Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection
Dim strSQL As String

On Error GoTo ErrorHandler:

Set cnn = CurrentProject.connection

'DoCmd.RunSQL "DELETE * FROM tblfrmCreateItemSubCount"
'DoCmd.OpenQuery "qryfrmCreateItemSubCount"
'strSQL = "SELECT * FROM tblfrmCreateItemSubCount"
strSQL = "SElECT NumSubPositions, NumSubFillers, NumCopperUnits FROM tblCableConstructionReferences "
strSQL = strSQL & " WHERE Base = '" & Forms!frmCreateItem!Base & "'"
rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic

If IsNull(rst!NumSubFillers) = True Then
    intNumSubFillers = 0
Else
    intNumSubFillers = rst!NumSubFillers
End If

If IsNull(rst!NumCopperUnits) = True Then
    intNumCopperUnits = 0
Else
    intNumCopperUnits = rst!NumCopperUnits
End If



intNumSub = rst!NumSubPositions



'Count the number of subs with active fiber
intUnitSeriesSubCount = intNumSub - intNumSubFillers - intNumCopperUnits
Forms!frmCreateItem!SubNo = intUnitSeriesSubCount
CountSubs = intUnitSeriesSubCount



rst.Close
cnn.Close


ErrorExit:
Exit Function

ErrorHandler:
rst.Close
cnn.Close
MsgBox ("Missing Number of Subs in Product Descriptions table")
Resume ErrorExit


End Function



Function CountSubBases() As Integer

Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection
Dim strSQL As String

On Error GoTo ErrorHandler:

Set cnn = CurrentProject.connection



strSQL = "SELECT * FROM qrySubunitBase WHERE Base = '" & Forms!frmCreateItem!Base & "'"
rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic

'Debug.Print rst.RecordCount


If rst.RecordCount <= 1 Then
    CountSubBases = 0
Else
    CountSubBases = 1
End If

'Debug.Print CountSubBases

ErrorExit:
Exit Function

ErrorHandler:
rst.Close
cnn.Close
Resume ErrorExit

End Function
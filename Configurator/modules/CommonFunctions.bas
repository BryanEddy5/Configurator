Option Compare Database
Option Explicit
Global sCutSheetItem As String
Global sFiberExtension As String
Global Const twips As Integer = 1440
Global bFurcation As Boolean
Global bCenterMember As Boolean

Function DeleteARecord()
    If MsgBox("Are you SURE you want to delete this record?" & vbCrLf & _
              "This will permanently delete the record.", vbYesNo, "2nd Delete Confirmation") = vbYes Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
End Function

'Function EmailAttachment(Attachment As String, EmailOption As Integer, Subject As String)
''
''
''    Dim oApp As New Outlook.Application
''    Dim oEmail As Outlook.MailItem
''    Dim fileName As String, todayDate As String
''    Dim StdOp As String
''    Dim rst As New ADODB.Recordset
''    Dim cnn As ADODB.Connection
''
''    Set cnn = CurrentProject.Connection
''
''    rst.Open "SELECT * FROM tblEmailOptions WHERE EmailOption = " & EmailOption & "", cnn, adOpenKeyset, adLockOptimistic
''
''
''
''    'Email the results of the report generated
''    Set oEmail = oApp.CreateItem(olMailItem)
''    With oEmail
''        .Recipients.Add rst!EmailTo
''        .Subject = rst!EmailSubject & " - " & Subject
''        .Body = rst!EmailBody & StdOp & vbCrLf & vbCrLf & _
''                "This is an automated email"
''        .Attachments.Add Attachment
''        .Send
''    End With
''
''    MsgBox "Email successfully sent!", vbInformation, "EMAIL STATUS"
'
'
'End Function

Function CreatePDF(Object As String, fileName As String, FilterCriteria As String) As String

    Dim todayDate As String

    todayDate = Format(Date, "MMDDYYYY")

    fileName = fileName & "_" & todayDate & ".pdf"
    DoCmd.OpenReport Object, acViewPreview, , FilterCriteria, acWindowNormal
    DoCmd.OutputTo acReport, Object, acFormatPDF, fileName, False
    DoCmd.Close acReport, Object
    
    CreatePDF = fileName


End Function

Sub FilesAndFolders()
    Dim FSO As Object                            'FileSystemObject
    Dim fldStart As Object                       'Folder
    Dim fld As Object                            'Folder
    Dim Mask As String
    
    On Error GoTo Error_Resume

    Set FSO = CreateObject("scripting.FileSystemObject") ' late binding
    'Set fso = New FileSystemObject 'or use early binding (also replace Object types)

    Set fldStart = FSO.GetFolder(modImages.ImagesCablePath) ' <-- use your FileDialog code here

    Mask = "*.jpg"
    Debug.Print fldStart.path & "\"
    ListFiles fldStart, Mask
    For Each fld In fldStart.SubFolders
        ListFiles fld, Mask
        ListFolders fld, Mask
    Next
    
Error_Resume:
    Resume Next
End Sub

Sub ListFolders(fldStart As Object, Mask As String)
    Dim fld As Object                            'Folder
    For Each fld In fldStart.SubFolders
        'If fld.Name Like "*cross*" Then
        'Debug.Print fld.Path & "\"
        ListFiles fld, Mask
        ListFolders fld, Mask
        'End If
    Next

End Sub

Sub ListFiles(fld As Object, Mask As String)

    Dim fl As Object                             'File
    Dim NewPath As String
    
    On Error GoTo Error_Resume
    
    NewPath = ImagesCablePath
    
    For Each fl In fld.Files
        If fl.name Like Mask Then
            'fl.Name = Left(fl.Name, 7) & "-" & Mid(fl.Name, 8, 5)
            'fl.Name = RenameImageFiber(fl.Name)
            'FileCopy fld.Path & "\" & fl.Name, NewPath & "\" & fl.Name     'Moving Files
            'fl.Name = Replace(fl.Name, "UX", "CX")                            'Renaming Files
            'Debug.Print fld.Path & "\" & fl.Name                             'File Discovery with directory
            
            'fl.Name = RenameImageFiber(fl.Name)
            'Debug.Print Replace(fl.Name, "Q", "X")
            'fl.Name = Replace(fl.Name, "Q", "X")
            'Debug.Print fld.Path & "\" & fl.Name & vbCrLf & modImages.ImagesCablePath & GetNewImageNameFromFileName(fl.Name, fld.Path)
            FileCopy fld.path & "\" & fl.name, ImagesCablePath & GetNewImageNameFromFileName(fl.name, fld.path)
            'RenameImageFiber (fl.Name)
        End If
    Next
    
Error_Resume:
    Resume Next
End Sub

Function BomComponentExplode(item As String) As String



    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim strStoredProcedure As String
    Dim rstTable As New ADODB.Recordset
    Dim cnnTable As New ADODB.connection



    Set cnnTable = CurrentProject.connection

    DoCmd.RunSQL "DELETE FROM tblBOMComponentExplode"


    If item Like "*-??bx*" Then
        item = Left(item, 9) & "0"
    End If

    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr



    strStoredProcedure = "  DECLARE @return_value int  EXEC    @return_value = [dbo].[BOMExplode_Component_Level]  N'" & item & "' "

    rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic

    rstTable.Open "tblBOMComponentExplode", cnnTable, adOpenKeyset, adLockOptimistic

    Do Until rst.EOF
        With rstTable
            .AddNew
            !AssemblyItemNumber = rst!AssemblyItemNumber
            !BOMLevel = rst!BOMLevel
            !Weight_kg_km = rst!Weight_kg_km
            .Update
        End With
    
        rst.MoveNext
    
    Loop


    rst.Close
    rstTable.Close
    cnnTable.Close
    cnn.Close



 
End Function

Function BomCostExplode(item As String) As String



    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim strStoredProcedure As String
    Dim rstTable As New ADODB.Recordset
    Dim cnnTable As New ADODB.connection



    Set cnnTable = CurrentProject.connection

    DoCmd.RunSQL "DELETE FROM tblBOMComponentExplode"

    'If Item Like "WTC%" Or "DNT%" Then
    '    Exit Function
    'End If


    If item Like "*-??bx*" Then
        item = Left(item, 9) & "0"
    End If

    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr


    strStoredProcedure = "  DECLARE @return_value int  EXEC    @return_value = [dbo].[GetExplodedBOMCost]  N'" & item & "' "

    rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic

    rstTable.Open "tblBOMComponentExplode", cnnTable, adOpenKeyset, adLockOptimistic

    Do Until rst.EOF
        With rstTable
            .AddNew
            !AssemblyItemNumber = rst!AssemblyItemNumber
            !BOMLevel = rst!BOMLevel
            !Weight_kg_km = rst!Weight
            !Cost_Ratio = rst!Cost_Ratio
            !ItemCost_km = rst!ItemCost_km
            !FinishedGood = rst!FinishedGood
            !ComponentItemNumber = rst!ComponentItemNumber
            !Total_Cost_km = rst!Total_Cost_km
            .Update
        End With
    
        rst.MoveNext
    
    Loop


    rst.Close
    rstTable.Close
    cnnTable.Close
    cnn.Close



 
End Function

Function FindItemData(item As String, DesiredField As String) As String

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    Dim strSQL As String
    Set cnn = CurrentProject.connection


    strSQL = "SELECT   dbo_CableUnion.Base as Base, tblCableConstructionReferences.BaseID as BaseID, tblDesignTypes.DesignTypeID as DesignTypeID, " & DesiredField
    strSQL = strSQL & " FROM ((dbo_CableUnion INNER JOIN tblCableConstructionReferences ON dbo_CableUnion.Base = tblCableConstructionReferences.Base) "
    strSQL = strSQL & " INNER JOIN tblDesignTypes ON tblCableConstructionReferences.DesignTypeID = tblDesignTypes.DesignTypeID) "
    strSQL = strSQL & " INNER JOIN tblCableConstructions ON tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID  "
    strSQL = strSQL & " WHERE dbo_CableUnion.Oracle = '" & item & "'"

    rst.Open (strSQL), cnn, adOpenKeyset, adLockOptimistic

    FindItemData = rst(DesiredField).value



    rst.Close
    cnn.Close

End Function

Function TensileStrengthCalculator(item As String) As Integer


    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim CoreItem As String

    If Mid(item, 11, 2) = "IA" Then
        CoreItem = FindCoreCable(item)
    
        item = FindExistingBOMItem(CoreItem)
    Else
        item = FindExistingBOMItem(item)

    End If

    DoCmd.RunSQL "DELETE FROM tblBOMComponentExplode"

    item = FindExistingBOMItem(item)

    strSQL = FormatString("  SELECT  [dbo].usf_TensileStrengthCalculator( '{0}') as TensileStrength_Newton", item)

    Set rst = Core_DataAccess.GetAdodbRecordset(strSQL)

    TensileStrengthCalculator = rst!TensileStrength_Newton


    rst.Close


End Function

Function BaseTensileStrengthCalculator(Base As String) As Integer


    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim strSQL As String
    Dim item As String




    item = FindItemFromBase(Base)

    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr

    strSQL = "  SELECT  [dbo].usf_TensileStrengthCalculator( '" & item & "') as TensileStrength_Newton"

    rst.Open (strSQL), cnn, adOpenKeyset, adLockOptimistic


    BaseTensileStrengthCalculator = rst!TensileStrength_Newton


    rst.Close
    cnn.Close





End Function

Function CableRatingPrint(printLine As String, DesignCode As String)

    Dim sBurnListing As String
    Dim BurnListing() As String
    Dim line As Variant


    'Determine if the listing is MSHA
    If DesignCode Like "*msha*" Then

        sBurnListing = ExtractTextFromString("P-", printLine, " ")



    Else
        BurnListing = Split(printLine, " ")
    
        For Each line In BurnListing
            If line Like "*of*" Then
                If checkBurnListing(line) Then
                    sBurnListing = line
                    Exit For
                End If
            End If
    
        Next
    
        'sBurnListing = ExtractTextFromString("OF", PrintLine, " ")
    

    End If

    CableRatingPrint = sBurnListing


End Function

Function checkBurnListing(SafetyRating As Variant) As Boolean
    Dim rst As New ADODB.Recordset
    Dim cnn As New ADODB.connection
    Dim sql As String
    
    checkBurnListing = False
    
    cnn.Provider = "sqloledb"
    cnn.Open Core_SqlConnections.ConnectionString
    
    sql = "Select DISTINCT Rating FROM tblBurnListingRatingReference WHERE Rating = '" & SafetyRating & "' AND Rating <> 'NONE' "
    
    rst.Open (sql), cnn, adOpenStatic, adLockReadOnly
    
    
    If rst.EOF = False Then
    
        checkBurnListing = True

    End If
    


End Function

Function CableRatingPrintCanada(printLine As String, DesignCode As String)

    Dim sBurnListing


    'Determine if the listing is MSHA
    If Not printLine Like "*FT*" Then
        Exit Function
    End If


    
    sBurnListing = Nz(sBurnListing, "") & " " & ExtractTextFromString("FT", printLine, " ")
    

    CableRatingPrintCanada = sBurnListing


End Function

Function ExtractTextFromString(sText As String, sString As String, sStringEnd) As String

    On Error GoTo ErrorHandler:

    If Len(Trim(sString)) = 0 Then
        Exit Function
    End If
    If Right(sString, 1) <> " " Then sString = sString & " "

    ExtractTextFromString = Mid(sString, InStr(sString, sText), InStr(InStr(sString, sText), sString, sStringEnd) - InStr(sString, sText))

    Exit Function

ErrorHandler:
    ExtractTextFromString = "ERROR"
    Exit Function

End Function

Function ExtensionCountLookup(sIndicator As String) As Integer

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    Dim strSQL As String
    Set cnn = CurrentProject.connection

    strSQL = "Select ExtensionNumber FROM tblCableExtensionReference WHERE ExtensionIndicator = '" & sIndicator & "'"

    rst.Open (strSQL), cnn, adOpenForwardOnly, adLockReadOnly

    If rst.EOF = False Then


        ExtensionCountLookup = rst!ExtensionNumber

    Else

        MsgBox "Extension indicator not found", vbCritical
    End If

End Function

Function Dec2Bin(ByVal DecimalIn As Variant) As String
              
    DecimalIn = CDbl(CDec(DecimalIn))
    Do While DecimalIn <> 0                      'Check if input is not zero
        Dec2Bin = Format$(DecimalIn - 2 * Int(DecimalIn / 2)) & Dec2Bin 'Divide the input by two and determine if the value is a whole number. 0 for whole and 1 for not
        DecimalIn = Int(DecimalIn / 2)           'Divid number by two for the next iteration  and repeat the step above
    Loop

End Function

Function Bin2Hex(BinNum As String) As String
    Dim BinLen As Integer, i As Integer, Cuts As Integer, HexFactor As Integer, j As Integer, LenSegment As Integer, Remainder As Integer
    Dim HexNum As Variant, HexHolder As Variant
    Dim HexSegment As String, HexPosition As Integer
    Dim bFlag As Boolean

    
    BinLen = Len(BinNum)                         'Determine the length of the number
    Cuts = Int(BinLen / 4)                       'find the number of segments the binary number should be divided into for segments of 4
    If BinLen Mod 4 <> 0 Then                    ' if it does not evenly divide then add a segment division to house to < 4 portion at the beginning of the binary number
        Remainder = BinLen Mod 4
        Cuts = Cuts + 1
        bFlag = True
    End If
    
    For j = 1 To Cuts                            'For each segment of 4 binary indicators a loop will be performed
        If j = Cuts And bFlag = True Then
            HexSegment = Left(Right(BinNum, 4 * (j)), Remainder) 'Retreive only the remaining binary digits necessary
        Else
            HexSegment = Left(Right(BinNum, 4 * (j)), 4) 'Starting from the right it grabs the first 4 numbers
        End If
        LenSegment = Len(HexSegment)
        HexHolder = 0
        HexFactor = 1                            'The first position of the 4 binary indicators will be multiplied by 2^3
        
        For i = LenSegment To 1 Step -1          'Iterate through each position of the 4 binary indicators in the segment
            HexPosition = Int(Mid(HexSegment, i, 1)) * HexFactor 'Go through each position and multiply by 2^exponent. The exponent is dependent on the position, starting at 3 and dropping 1 for each subsequent position.
            HexHolder = HexHolder + HexPosition  'Add each positions value together
            HexFactor = HexFactor * 2            'drop the exponent by 1 for each postion 2^3, 2^2, 2^1, 2^1
    
        Next i
        If HexHolder > 9 Then
            HexHolder = HexLetter(HexHolder)     'Calls a function to find the letter if above 9
        End If
            
        HexNum = HexHolder & HexNum              'Places the HexNum in position
    Next j
    
    Bin2Hex = HexNum


End Function

Function HexLetter(Number As Variant) As String

    Select Case Number
    Case 1 To 9
        HexLetter = Number
    Case 10
        HexLetter = "A"
    Case 11
        HexLetter = "B"
    Case 12
        HexLetter = "C"
    Case 13
        HexLetter = "D"
    Case 14
        HexLetter = "E"
    Case 15
        HexLetter = "F"
    Case 16
        HexLetter = "G"
    Case 18
        HexLetter = "I"
    Case 24
        HexLetter = "O"
    Case 32
        HexLetter = "P"
    Case 36
        HexLetter = "Q"
    Case 48
        HexLetter = "R"
    Case 60
        HexLetter = "S"
    Case 72
        HexLetter = "T"
    End Select
End Function

Function FindStandards(BaseID As Integer) As String

    Dim strSQL As String
    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection

    Set cnn = CurrentProject.connection

    strSQL = "SELECT Standard "
    strSQL = strSQL & " FROM tblCableConstructions INNER JOIN Standards_CableConstructionStandards ON tblCableConstructions.BaseID = Standards_CableConstructionStandards.BaseID "
    strSQL = strSQL & " WHERE Standards_CableConstructionStandards.BaseID = " & BaseID & " ORDER BY Standard"

    rst.Open (strSQL), cnn, adOpenStatic, adLockReadOnly

    Do Until rst.EOF
        FindStandards = rst!Standard & vbCrLf & FindStandards
        rst.MoveNext
    Loop






End Function

Function FindSafetyTest(BurnListing As String) As String


    Dim strSQL As String
    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection

    BurnListing = Trim(BurnListing)

    Set cnn = CurrentProject.connection

    strSQL = "SELECT Standard "
    strSQL = strSQL & " FROM tblStandardsSafetyListing "
    strSQL = strSQL & " WHERE SafetyListing = '" & BurnListing & "' ORDER BY Standard"

    rst.Open (strSQL), cnn, adOpenStatic, adLockReadOnly

    Do Until rst.EOF
        FindSafetyTest = rst!Standard & vbCrLf & FindSafetyTest
        rst.MoveNext
    Loop

End Function

Function AddEmailAddress(sUser As String, sUserEmail As String)

    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim provstr As String
    Dim strStoredProcedure As String
    Dim rst As New ADODB.Recordset



    On Error GoTo ErrorHandler:




    strStoredProcedure = "exec Premise.dbo.usp_UserEmailAddress  '" & sUser & "', '" & sUserEmail & "'"


    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr



    rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic
 


    cnn.Close

ErrorExit:
    Exit Function

ErrorHandler:
    MsgBox "An error occurred during submission.", vbCritical
    Resume ErrorExit


End Function

Function WhereUsed(item As String) As String



    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim strStoredProcedure As String
    Dim rstTable As New ADODB.Recordset
    Dim cnnTable As New ADODB.connection



    Set cnnTable = CurrentProject.connection

    DoCmd.RunSQL "DELETE FROM tblWhereUsed"


    If item Like "*-??bx*" Then
        item = Left(item, 9) & "0"
    End If

    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"
    cnn.CommandTimeout = 60


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr


    strStoredProcedure = "  DECLARE @return_value int  EXEC    @return_value = [dbo].[usp_WhereUsed]  N'" & item & "' "

    rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic

    rstTable.Open "tblWhereUsed", cnnTable, adOpenKeyset, adLockOptimistic

    Do Until rst.EOF
        With rstTable
            .AddNew
            !AssemblyItemNumber = rst!AssemblyItemNumber
            !ComponentItemNumber = rst!ComponentItemNumber
            !Quantity = rst!Quantity
            .Update
        End With
    
        rst.MoveNext
    
    Loop


    rst.Close
    rstTable.Close
    cnnTable.Close
    cnn.Close



 
End Function

Public Function RoundTensile(Weight As Double) As Double


    Select Case Weight
    Case Is < 10
        RoundTensile = Round(Weight, 1)
    Case 10 To 45
        RoundTensile = Round(Weight, 0)
    Case 45 To 100
        RoundTensile = Round(Weight / 5, 0) * 5
    Case Is > 100
        RoundTensile = Ceiling(Weight / 10) * 10
    
        
        
    End Select
        
        

End Function

Public Function RoundRanges(Weight As Double) As Double


    Select Case Weight
    Case Is < 10
        RoundRanges = Round(Weight, 1)
    Case 10 To 100
        RoundRanges = Round(Weight, 0)
    Case 100 To 200
        RoundRanges = Ceiling(Weight / 5) * 5
    Case Is > 200
        RoundRanges = Ceiling(Weight / 10) * 10
    End Select
        

End Function

Public Function Ceiling(ByVal x As Double, Optional ByVal factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(x / factor) - (x / factor - Int(x / factor) > 0)) * factor
End Function

Function GetSalesForcePricing(Base As String) As String



    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim strStoredProcedure As String
    Dim rstTable As New ADODB.Recordset
    Dim cnnTable As New ADODB.connection



    Set cnnTable = CurrentProject.connection

    DoCmd.RunSQL "DELETE FROM tblSalesForce_Catalog_Pricing"



    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"
    cnn.CommandTimeout = 60


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr



    strStoredProcedure = "  DECLARE @return_value int  EXEC    @return_value = [dbo].[usp_GetSalesForcePricing]  N'" & Base & "' "

    rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic

    rstTable.Open "tblSalesForce_Catalog_Pricing", cnnTable, adOpenKeyset, adLockOptimistic

    Do Until rst.EOF
        With rstTable
            .AddNew
            !Base = rst!Base
            !Price = rst!Price
            !Fiber_Type__c = rst!Fiber_Type__c
            !CPQ_Fiber_Cost = rst!CPQ_Fiber_Cost
            !FiberCount = rst!FiberCount
            !TotalFiberCost = rst!TotalFiberCost
            !Contribution = rst!Contribution
            !Multiplier = rst!Multiplier
            !LoadedBaseCost = rst!LoadedBaseCost
            !BaseSkeletalCost = rst!BaseSkeletalCost
            !Product_Category = rst!Product_Category
            !FiberID = rst!FiberID
            !ArmoredPrice = Nz(rst!ArmoredPrice, 0)
            !ArmorCostAdder = Nz(rst!ArmorCostAdder, 0)
            !Armor = Nz(rst!Armor)
            .Update
        End With
    
        rst.MoveNext
    
    Loop


    rst.Close
    rstTable.Close
    cnnTable.Close
    cnn.Close



 
End Function

Function DoesPIDExist(Customer As String, PID As String) As Boolean



    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim sql As String
    Dim cnnTable As New ADODB.connection



    Set cnnTable = CurrentProject.connection

    DoCmd.RunSQL "DELETE FROM tblNewProductConstructionHold"



    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"
    cnn.CommandTimeout = 60


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr


    sql = " Select [New Oracle Part #], Customer, [Customer Part#], [Item No], Base FROM [Basic Product Construction] WHERE [Customer Part#] <> 'CORE' AND Customer = '" & Customer & "' AND [Customer Part#] = '" & PID & "' AND Active <> 0"

    rst.Open (sql), cnn, adOpenKeyset, adLockOptimistic


    If rst.EOF = True Then
        rst.Close
        Exit Function
    End If

    DoesPIDExist = True




    rst.Close
    cnn.Close


End Function

Function ExtractNumber(sText As String)
    Dim iCount As Integer, i As Integer
    Dim lNum As String
     
    For iCount = Len(sText) To 1 Step -1
        If IsNumeric(Mid(sText, iCount, 1)) Then
            i = i + 1
            lNum = Mid(sText, iCount, 1) & lNum
        End If
        If i = 1 Then lNum = CInt(Mid(lNum, 1, 1))
    Next iCount
     
    ExtractNumber = CLng(lNum)
End Function

Public Function Load() As Boolean

    Dim cdb As DAO.Database
    Dim rst As DAO.Recordset
    Dim qdf As DAO.QueryDef

    Set cdb = CurrentDb

    On Error GoTo HandleError:

    'Dim qdf As QueryDef
    'Set qdf = CurrentDb.QueryDefs("tblItemPassToSQL")
    'qdf.Connect = CurrentDb.TableDefs("tblMaterialBOMExplode").Connect
    'qdf.sql = " INSERT INTO tblItemPassToSQL SELECT * FROM tblMaterialDescriptions "
    'qdf.Execute dbFailOnError
    'qdf.Close
    'Set qdf = Nothing

    DoCmd.OpenQuery "tblItemPassToSQL"


    Set rst = cdb.OpenRecordset("tblMaterialBOMExplode", dbOpenSnapshot)
    Do While Not rst.EOF
        Set qdf = cdb.CreateQueryDef("", "INSERT INTO [##ItemPassToSQL] (Item,Quantity) VALUES ('" & rst!OracleItem & "', " & rst!Quantity & ")")
        qdf.connect = "ODBC;DRIVER=SQL Server;SERVER=NAASPB-PRD04\SQL2014;UID=eddybc;Trusted_Connection=Yes;DATABASE=Premise;Network=DBMSSOCN"
        qdf.ReturnsRecords = False
        qdf.Execute dbFailOnError
        Set qdf = Nothing
        rst.MoveNext
    Loop

    RetrieveData

    rst.Close
    Set rst = Nothing


Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading data:" & vbCrLf & Err.Description, vbCritical
    Resume Done
    '
End Function

Public Function RetrieveData() As Boolean


    On Error GoTo HandleError:

    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim sQry As String
    Dim rst As New ADODB.Recordset


    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"

    ' Specify connection string on Open method.
    cnn.Open "Server=NAASPB-PRD04\SQL2014;Database=tempdb;Trusted_Connection=yes"
    sQry = "SELECT * FROM #ItemPassToSQL"

    rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
    RetrieveData = False


    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If
    
        ' Note use of Nz function in case of Null values in the table
        'mlApplicationDirectory = Nz(!ApplicationDirectory)

        .Close
    End With
    RetrieveData = True

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
    '
End Function

Function TestTempTable()
    '-- Reference Library
    '-- Microsoft ActiveX data objects 2.6 library needed for ADO
    '-- Reference Library
    '-- Microsoft ActiveX data objects 2.6 library needed for ADO
    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rs As New ADODB.Recordset

    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"

    ' Specify connection string on Open method.
    cnn.Open Core_SqlConnections.ConnectionString


    '-- create the temporary table
    Dim mykey As Integer, sql1 As String, textData As String
    textData = "Some description data"
    mykey = 20
    sql1 = "insert into ##myTempTable values(" & mykey & ",'" & textData & "')"
    cnn.Execute "create table ##myTempTable (akey int, muchWords ntext) "
    Debug.Print sql1
    cnn.Execute sql1
    rs.Open "Select * from ##myTempTable", cnn, adOpenStatic, adLockOptimistic
    Debug.Print "first field = "; rs(0); " second field = "; rs(1)

    ''cn.Execute "drop table #myTemptable"
    cnn.Close
    Set cnn = Nothing

End Function

Public Function LoadData() As Boolean


    Dim rsMaterial As New ADODB.Recordset
    Dim cnnMaterial As New ADODB.connection
    Set cnnMaterial = CurrentProject.connection
    Dim strSQL As String

    strSQL = "Select * from tblMaterialBOMExplode"

    rsMaterial.Open (strSQL), cnnMaterial, adOpenKeyset, adLockOptimistic


    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim sExec As String, OracleItem As String
    Dim Quantity As Integer
    Dim sql1 As String

    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"

    ' Specify connection string on Open method.
    cnn.Open Core_SqlConnections.ConnectionString




    sExec = " IF OBJECT_ID(N'tempdb..##ItemPassToSQL', N'U') IS NOT NULL "
    sExec = sExec & " DROP TABLE ##ItemPassToSQL "
    sExec = sExec & " CREATE TABLE ##ItemPassToSQL(OracleItem nvarchar(50),Quantity float,) "

    cnn.Execute sExec

    With rsMaterial

        Do Until .EOF
        
            OracleItem = Nz(!OracleItem)
            Quantity = Nz(!Quantity)
    
            sql1 = FormatString("insert into ##ItemPassToSQL values('{0}',{1})", OracleItem, Quantity)
        
            'Debug.Print sql1
            cnn.Execute sql1
            .MoveNext
        Loop

    End With

    'rs.Open "Select * from ##ItemPassToSQL", cnn, adOpenStatic, adLockOptimistic
    '
    'Do Until rs.EOF
    '    Debug.Print "first field = "; rs(0); " second field = "; rs(1)
    '    rs.MoveNext
    'Loop
    '


    Call GetBuyMaterials

    ''cn.Execute "drop table #myTemptable"
    cnn.Close
    Set cnn = Nothing




Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading data:" & vbCrLf & Err.Description, vbCritical
    Resume Done
    '
End Function

Public Function GetBuyMaterials() As Boolean

    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim strStoredProcedure As String
    Dim rstTable As New ADODB.Recordset
    Dim cnnTable As New ADODB.connection


    Set cnnTable = CurrentProject.connection

    DoCmd.RunSQL "DELETE FROM tblGetBuyMaterials"



    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr


    strStoredProcedure = "  DECLARE @return_value int  EXEC    @return_value = [dbo].[usp_GetBuyMaterials] "

    rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic

    rstTable.Open "tblGetBuyMaterials", cnnTable, adOpenKeyset, adLockOptimistic

    Do Until rst.EOF
        With rstTable
            .AddNew
            !OracleItem = Nz(rst!OracleItem)
            !Quantity = Nz(rst!Quantity)
            !ComponentItemNumber = Nz(rst!ComponentItemNumber)
            !ComponentQuantity_PerQty = Nz(rst!ComponentQuantity_PerQty)
            !TotalCost = Nz(rst!TotalCost)
            !PrimaryUOM = Nz(rst!PrimaryUOM)
            !ComponenetQuantity_M = Nz(rst!ComponenetQuantity_M)
            !StartupScrap = Nz(rst!StartupScrap)
            .Update
        End With
    
        rst.MoveNext
    
    Loop


    rst.Close
    rstTable.Close
    cnnTable.Close
    cnn.Close



    '
End Function

Function SignificantFigure(Number As Double, SignificantFigures As Integer)

    'SignificantFigure = Round(Number, SignificantFigures - Int(Log10(Abs(Number))))

    SignificantFigure = CLng(Number * (10 ^ SignificantFigures)) / (10 ^ SignificantFigures)

End Function

Static Function Log10(x)
    Log10 = Log(x) / Log(10#)
End Function

Public Function getColorChip(Color As String, JacketMaterial As String) As Variant

    Dim rst As New ADODB.Recordset
    Dim sql As String

    sql = "SELECT ColorChip FROM dbo.vColor_Chip_Jacket_Material where Color = '" & Color & "' AND JacketMaterial = '" & JacketMaterial & "'"

    Set rst = Core_DataAccess.GetAdodbRecordset(sql)


    If rst.EOF Then
        getColorChip = Null
        Exit Function
    End If


    With rst
        getColorChip = Nz(!ColorChip)
    End With
    
    
    rst.Close

End Function

Public Sub DeleteFile(ByVal FileToDelete As String)
    If CommonFunctions.FileExists(FileToDelete) Then 'See above
        ' First remove readonly attribute, if set
        SetAttr FileToDelete, vbNormal
        ' Then delete the file
        Kill FileToDelete
    End If
End Sub

Public Function FileExists(ByVal FileToTest As String) As Boolean
    FileExists = (Dir(FileToTest) <> "")
End Function

Public Sub subRefreshDesktop()
    'Dim objshell As Object
    'Dim objDesktop As Object
    '
    'Set objshell = CreateObject("Shell.Application")
    'Set objDesktop = objshell.NameSpace(0)
    'objDesktop.Self.InvokeVerb "R&efresh"

    Dim WSHShell As Object
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.AppActivate "Program Manager"
    WSHShell.SendKeys "{F5}"
    Set WSHShell = Nothing


End Sub

Public Function getColorByColorLetter(ColorLetter$) As Variant

    On Error GoTo HandleError:

    Dim sQry As String
    Dim rst As New ADODB.Recordset
    Dim Color$
    sQry = "SELECT * FROM tblColors_Available WHERE Color_ID = '" & ColorLetter & "'"


    Set rst = Core_DataAccess.GetAdodbRecordset(sQry)

    With rst
        If .EOF Then
            getColorByColorLetter = Null
            GoTo Done
        End If
        
        getColorByColorLetter = !Color

        .Close
    End With

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while getting color:" & vbCrLf & Err.Description, vbCritical
    Resume Done
    '
End Function

Public Function UnitSeriesExist(UnitSeries$) As Boolean


    On Error GoTo HandleError:

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
    sQry = "SELECT * FROM [dbo].[CableUnion] WHERE [Unit Series] = '" & UnitSeries & "'"


    rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic



    With rst
        If .EOF Then
            UnitSeriesExist = False
            GoTo Done
        End If
        
        UnitSeriesExist = True
        
        ' Note use of Nz function in case of Null values in the table

        .Close
    End With

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " checking if Unit Series exists:" & vbCrLf & Err.Description, vbCritical
    Resume Done
    '
End Function

Public Function GetFileName(path$) As String

    GetFileName = Right(path, Len(path) - InStrRev(path, "\"))

End Function


Public Function FormatString(ParamArray arr() As Variant) As String

    Dim i As Long
    Dim temp As String

    temp = CStr(arr(0))
    For i = 1 To UBound(arr)
        temp = Replace(temp, "{" & i - 1 & "}", CStr(arr(i)))
    Next

    FormatString = temp
End Function
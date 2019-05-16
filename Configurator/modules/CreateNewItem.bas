'@NoIndent
Option Compare Database
Option Explicit
Public CableQualities As New clsCableQualities
Private m_ItemRequestNumber
Private mlDesignExtension

Public Property Let setIsPrinted(IsPrinted As Boolean)
    CableQualities.IsPrinted = IsPrinted

End Property

Public Property Get ItemRequestNumber() As String
    ItemRequestNumber = m_ItemRequestNumber

End Property

Public Property Let ItemRequestNumber(ItemRequestNumber$)

    m_ItemRequestNumber = ItemRequestNumber

End Property

Public Property Let NewItemNumber(NewItemNumber$)
    CableQualities.ItemNumber = NewItemNumber

End Property

Public Property Get NewItemNumber() As String
    NewItemNumber = CableQualities.ItemNumber

End Property

Public Property Get getIsPrinted() As Boolean
    getIsPrinted = CableQualities.IsPrinted

End Property

Public Function GetisVendorSpecificPrint() As Boolean
    GetisVendorSpecificPrint = mFiberPrint.isVendorSpecificPrint

End Function

Public Function CreateNewItemIsPrinted() As Boolean

    CreateNewItemIsPrinted = CableQualities.IsPrinted

End Function

Public Function getNewItemListingLabel() As String

    getNewItemListingLabel = CableQualities.getListingLabel

End Function

Function CreateNewItemPrintSpacing() As String

End Function

Function CreateNewItemNewPrintLine4() As String


End Function

Function getJacketColor(CableColorIndicator As String) As String

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    Dim strColorID As String


    Set cnn = CurrentProject.connection
    rst.Open "SELECT * FROM tblColors_Available WHERE Color_ID = '" & CableColorIndicator & "' ORDER BY Color DESC", cnn, adOpenKeyset, adLockOptimistic
    'Me.ColorChip = "NONE"
    'Me.Color.RowSourceType = "Value List"
    getJacketColor = rst!Color
    'rst.Update
    rst.Close
    cnn.Close


End Function

Function FindFurcation(Base As String, Location As Integer, FurcationTube As String) As Boolean
    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim strSQL As String

    If Location = 2 Then
        FindFurcation = True
        Exit Function
    End If

    Set cnn = CurrentProject.connection

    strSQL = "SELECT * FROM (tblFurcationTubeReference as tf INNER JOIN tblDesignCodeFurcationTube as tdf ON tf.Furcation = tdf.Furcation) "
    strSQL = strSQL & " INNER JOIN tblCableConstructionReferences tblfurc ON tf.BaseID = tblfurc.BaseID "
    strSQL = strSQL & " WHERE Base = '" & Base & "' AND FurcationLetter = '" & FurcationTube & "'"

    rst.Open strSQL, cnn, adOpenKeyset, adLockBatchOptimistic

    If rst.RecordCount = 0 Then
        FindFurcation = False
        MsgBox ("Cable construction does not allow for this Furcation tube"), vbCritical
    Else
        FindFurcation = True
    End If

    rst.Close
    cnn.Close

End Function

Function CatalogExist(Customer As String, PID As String, Base As String, Fiber1 As String, Fiber2 As String, DesignCode As String, RibbonHighCure As Boolean) As Boolean



    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim sql As String
    Dim cnnTable As New ADODB.connection
    Dim search As String




    Set cnnTable = CurrentProject.connection


    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"
    cnn.CommandTimeout = 60


    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr

    search = "[Fiber2] = '" & Fiber2 & "' AND [Fiber] = '" & Fiber1 & "' AND [Base] ='" & Base _
           & "' AND RibbonHighCure = " & CInt(Int(RibbonHighCure)) & "  AND ItemNo = '" & DesignCode _
           & "' AND Customer = '" & Customer & "' AND PID = '" & PID & "'"

    sql = " Select Oracle FROM vItemSearch WHERE PID <> 'CORE' AND " & search

    rst.Open (sql), cnn, adOpenKeyset, adLockOptimistic

    CatalogExist = False


    If rst.EOF = True Then
        rst.Close
        Exit Function
    End If

    CatalogExist = True

       
    DoCmd.OpenForm "frmFindCustomerPID", acNormal, , search
        
    MsgBox " This may be an existing item.  Review the items in the pop up.", vbCritical


    rst.Close
    cnn.Close


End Function

Public Function ItemExtensionDesignCheck(DesignCode$) As Boolean
    ''Iterate through the design code and possible delimiters to determine if there is a match for the design code extension''
    Dim Extension$(), hold$()
    Dim i As Integer, delimiter, j
    Dim deliminate$(2)

    ItemExtensionDesignCheck = False

    deliminate(0) = ":"
    deliminate(1) = "-"
    hold = Split(DesignCode, ":")
    For Each j In hold
        Extension = Split(j, "-")
    Next
    For Each delimiter In deliminate
        For i = 0 To UBound(Extension)
            If CheckDesignExtensionExists(Extension(i), delimiter) = True Then
                ItemExtensionDesignCheck = True
                Exit Function
            End If
    
        Next

    Next

End Function

Private Function CheckDesignExtensionExists(Extension$, delimiter) As Boolean
    ''Check if extension matches existing design code extensions''

    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim sql As String

    CheckDesignExtensionExists = False

    cnn.Provider = "sqloledb"
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr


    sql = "Select DesignExtension FROM vCosntructionReferenceActiveDesignCodeExtensions WHERE DesignExtension = '" & delimiter & Extension & "'"

    rst.Open (sql), cnn, adOpenKeyset, adLockOptimistic

    If Not rst.EOF Then
        mlDesignExtension = rst!DesignExtension
        CheckDesignExtensionExists = True
    End If


End Function

Function GetDesignCodeExtension() As String

    GetDesignCodeExtension = mlDesignExtension
    
End Function

Public Function UpdateToUnprintedCatalog(Catalog$) As String
    Dim hold$(), i
    hold = Split(Catalog, "-")

    For Each i In hold
        If i = hold(0) Then
            UpdateToUnprintedCatalog = Left(i, 8) + Replace(i, Mid(i, 9, 1), "U", 9, 1)
        ElseIf Not i Like "*box*" And Not i Like "*msha*" Then
            UpdateToUnprintedCatalog = UpdateToUnprintedCatalog + "-" + i
        End If

    Next


End Function

Function InsertRequestToLog() As Boolean



    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim provstr As String
    Dim sql As String

    If CableQualities.ItemNumber = "" Or IsNull(ItemRequestNumber) Then Exit Function


    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"



    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr


    sql = "SELECT * FROM dbo.tblItemBuildDescrepancyLog WHERE Item = '" & NewItemNumber & "'"

    rst.Open (sql), cnn, adOpenKeyset, adLockOptimistic

    InsertRequestToLog = False

    If Not rst.EOF Then
        MsgBox "This item already exists in the Item Build Log."
        Exit Function
    End If



    With rst
        .AddNew
        !ItemRequest = ItemRequestNumber
        !item = NewItemNumber
        .Update
    End With

    InsertRequestToLog = True
    rst.Close
    cnn.Close


End Function

Public Sub ArmorSelection(CatalogCode As String)

    Forms!frmCreateItem!Location = 2

    '    If CatalogCode Like "l*" Then
    '        Call LooseTube_Armor(CatalogCode)
    '        Exit Sub
    '    End If
    '
    If CatalogCode Like "*AIAP*" Then
        Forms!frmCreateItem!ArmorType = 1
        Forms!frmCreateItem!ListingCompany = "ETL"
        Forms!frmCreateItem!Jacket = "PVC00130"
        
    ElseIf CatalogCode Like "*AIAR*" And CatalogCode Like "*msha*" Then
        Forms!frmCreateItem!ArmorType = 7
        Forms!frmCreateItem!ListingCompany = "MSHA"
        Forms!frmCreateItem!Jacket = "PVC00117"
        Forms!frmCreateItem!PrintLine3 = "CAUTION! Not to be used where permissible equipment is required"

        
    ElseIf CatalogCode Like "*AIAR*" Then
        Forms!frmCreateItem!ArmorType = 3
        Forms!frmCreateItem!ListingCompany = "ETL"
        Forms!frmCreateItem!Jacket = "PVC00145"
                       
    
        
    ElseIf CatalogCode Like "*AIAE*" Then
        Forms!frmCreateItem!ArmorType = 4
        Forms!frmCreateItem!ListingCompany = "NONE"
        
        
    ElseIf CatalogCode Like "*AIAL*" Then
        Forms!frmCreateItem!ArmorType = 5
        Forms!frmCreateItem!Jacket = "POL00127"
        Forms!frmCreateItem!PrintLine3 = "115304 (ETL) TYPE OFCR-LS c(ETL) RoHS"
        Forms!frmCreateItem!ListingCompany = "ETL"
    
    End If

    Forms!frmCreateItem!cbSubunit.Enabled = False


End Sub

Public Function LooseTube_Armor(CatalogCode As String) As Boolean

    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim dblCableOD As Double
    Dim JacketType As Integer
    
    Set cnn = CurrentProject.connection
    
    dblCableOD = InputBox("What is the cable's OD?", "Cable OD")
    
    Select Case Left(CatalogCode, 2)
    Case "LL"
        JacketType = 6
    Case "LV"
        JacketType = 3
    Case Else
        'MsgBox "Cannot find armor type"
        Exit Function
    End Select
    
    rst.Open "SELECT * FROM tblArmoredSetups WHERE JacketType =  " & JacketType & " AND MaxCoreOD >= " & dblCableOD & " ORDER BY MaxCoreOD ASC", cnn, adOpenKeyset, adLockOptimistic
    
    If rst.EOF Then
        Exit Function
    End If
    
    LooseTube_Armor = True
    
    Forms!frmMainMenu!Bypass = -1
    
    Forms!frmCreateItem!ArmorNo = rst!ArmorNo
    Forms!frmCreateItem!ListingCompany = "ETL"
    Forms!frmCreateItem!Jacket = rst!JacketMaterial
    Forms!frmCreateItem!Base = rst!OracleBase
    Forms!frmCreateItem!ArmorCoreOD = rst!NomOD
    Forms!frmCreateItem!NomOD = rst!MaxJacketOD
    
    Forms!frmCreateItem!PrintType = 2
    Forms!frmCreateItem!Armor = -1
    Forms!frmCreateItem![Exists].value = 0
    Forms!frmCreateItem!btnUnitSeries.Visible = False
    Forms!frmCreateItem!Base.Enabled = False
    Forms!frmCreateItem!Jacket.Enabled = False
    Forms!frmCreateItem!cbSubunit.Enabled = False
    
    Forms!frmCreateItem![NewOracle#] = Forms!frmCreateItem![NewOracle#] & "IA"
    
    
        
    Forms!frmCreateItem!Color.rowSource = "qryfrmCreateItemChipColor_LooseTubeArmor"
    
    

End Function
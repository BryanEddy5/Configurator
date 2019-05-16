Option Compare Database
Option Explicit

Public Const schema_field_types = "1;dbBoolean;2;dbByte;3;dbInteger;4;dbLong;5;dbCurrency;6;dbSingle;7;dbDouble;8;dbDate;9;dbBinary;10;dbText;11;dbLongBinary;12;dbMemo;15;dbGUID;16;dbBigInt;17;dbVarBinary;18;dbChar;19;dbNumeric;20;dbDecimal;21;dbFloat;22;dbTime;23;dbTimestamp"
Public glb_coll As New Collection           ' used to xfer args to report

Public Const acDispControlType_CheckBox = 106
Public Const acDispControlType_TextBox = 109
Public Const acDispControlType_ListBox = 110
Public Const acDispControlType_ComboBox = 111

Public Const fig_TwipsPerInch = 1440
Public Const fig_TwipsPerCm = 567
Public Const figConnectIntro = ";DATABASE="
'

Public Sub DBM_Document_Tables(documentTableName$ _
                    , Optional tgtDb$ = "" _
                    , Optional workTabInCurrent As Boolean = True _
                    , Optional skipReport As Boolean = True _
                    , Optional skipTabPrp As Boolean = False _
                    , Optional skipTabIdx As Boolean = False _
                    , Optional skipRelations As Boolean = False _
                    )
    
'*******************************************
'Purpose:   create Schema table with data definitions for each of the tables in a d/b
'           optionally, produce report
'Author:    CD
'Date:      april 1999
'Called by: Free-standing rtn, called manually
'*******************************************
'   documentTableName$      name of (new) table in which to store Schema
'   tgtDb$                  name of db to be analysed - "" = current db
'   workTabInCurrent        is "documentTableName$" (to be) in currentdb or target db
'                           ... only meaningful if tgtDb$ <> ""
'   skipReport              don't run report
'   skipTabPrp              don't analyse table Properties
'   skipTabIdx              don't analyse table Indexes
'   skipRelations           don't analyse inter-table relationships

'   The FieldSeq field in the output Schema table is structured as:
'       Relationships   from 1 to ...
'       Tbl Properties  from 1 to 120
'       Indexes         from 121 to ...
'       Fields          from 2001 to ...

Const seqBase_prp = 0
Const seqBase_idx = 120
Const seqBase_fld = 2000

    Dim DBtgt As DAO.Database, DBworkTab As DAO.Database
    Dim rs As DAO.Recordset, tbD As DAO.TableDef, fld As DAO.Field, prp As Property
    Dim indx As DAO.Index
    Dim sql$, dTablName$, seq%, pseq%, skipTab As Boolean, wk$, tblName$
    Dim rsTst As DAO.Recordset
    Dim rel As DAO.Relation
    
    Dim durProc As Long, endProc As Date, begProc As Date
    
    dTablName$ = IIf(Left$(documentTableName$, 1) <> "[", _
                        documentTableName$, _
                        Mid$(documentTableName$, 2, Len(documentTableName$) - 2))
    
    If tgtDb$ = "" Then
        tgtDb$ = CurrentDb.name
        End If
    Set DBtgt = OpenDatabase(tgtDb$, True, False)
    
    If workTabInCurrent Then
        Set DBworkTab = CurrentDb
        Else
        Set DBworkTab = DBtgt
        End If
    
    begProc = Now
    BeginTrans
    On Error Resume Next
    Set tbD = DBworkTab.TableDefs(dTablName$)
    If Err Then
        Err.Clear
        Set tbD = DBworkTab.CreateTableDef(dTablName$)
        tbD.Fields.Append tbD.CreateField("TableName", dbText, 250)
        tbD.Fields.Append tbD.CreateField("FieldName", dbText, 250)
        tbD.Fields.Append tbD.CreateField("FieldType", dbText, 250)
        tbD.Fields.Append tbD.CreateField("FieldSize", dbText, 250)
        tbD.Fields.Append tbD.CreateField("FieldAllowZ", dbBoolean)
        tbD.Fields.Append tbD.CreateField("FieldPurpose", dbText, 250)
        tbD.Fields.Append tbD.CreateField("FieldComment", dbText, 250)
        tbD.Fields.Append tbD.CreateField("FieldSeq", dbInteger)
        Set fld = tbD.Fields("FieldPurpose"):        fld.AllowZeroLength = True
        Set fld = tbD.Fields("FieldComment"):        fld.AllowZeroLength = True
        DBworkTab.TableDefs.Append tbD
        
        DBM_EditTable_FieldLookup DBworkTab.name, dTablName$, "FieldType", schema_field_types, "Value List"
'        Set tbD = DBtgt.TableDefs(dtablName$)
'        Set fld = tbD.Fields("FieldPurpose"):        fld.AllowZeroLength = True
'        Set fld = tbD.Fields("FieldComment"):        fld.AllowZeroLength = True
        Else
        Err.Clear
        sql$ = "Delete * from [" & dTablName$ & "]"
        DBworkTab.Execute sql$
        End If
    Set rs = DBworkTab.OpenRecordset(dTablName$, dbOpenDynaset)
    If Not skipRelations Then
        seq = seqBase_prp
        For Each rel In DBtgt.Relations
            rs.AddNew
            seq = seq + 1
            rs!TableName = " Relation: " & rel.name
            rs!FieldName = rel.Table & ":: " & rel.foreignTable
            rs!FieldType = rel.Attributes
            wk = ""
            wk = wk & IIf(rel.Attributes And dbRelationUnique, "One-to-One; ", "One-to-Many; ")
            wk = wk & IIf(rel.Attributes And dbRelationDontEnforce, "No RefInteg; ", "Using RefInteg; ")
            wk = wk & IIf(rel.Attributes And dbRelationInherited, "RelationInherited; ", "")
            wk = wk & IIf(rel.Attributes And dbRelationUpdateCascade, "Cascade Updates; ", "")
            wk = wk & IIf(rel.Attributes And dbRelationDeleteCascade, "Cascade Deletes; ", "")
            '       0 = (not defined)   'The relationship is one-to-many, uses RI.
            '       1 = dbRelationUnique        'The relationship is one-to-one.
            '       2 = dbRelationDontEnforce   'The relationship isn't enforced (no referential integrity).
            '       4 = dbRelationInherited     'The relationship exists in a non-current database that contains the two linked tables.
            '     256 = dbRelationUpdateCascade 'Updates will cascade.
            '    4096 = dbRelationDeleteCascade 'Deletions will cascade.
            '16777216 = dbRelationLeft          'Microsoft Access only. In Design view, display a LEFT JOIN as the default join type.
            '33554432 = dbRelationRight         'Microsoft Access only. In Design view, display a RIGHT JOIN as the default join type.
            rs!FieldComment = wk

            wk = ""
            For Each fld In rel.Fields
                wk = wk & fld.name & ":: " & fld.ForeignName & "; "
                Next fld
            rs!FieldPurpose = wk
            rs!FieldSeq = seq       ' reporting sequence
            rs.Update
            Next rel
        End If
    For Each tbD In DBtgt.TableDefs
        tblName$ = tbD.name
        If tblName$ = dTablName$ Then
            skipTab = True
            Else
            skipTab = Exclude_based_on_Name(tblName$)
            End If
        If Not skipTab Then
            ' OK so far ... try to access table
            Err.Clear
            Set rsTst = DBtgt.OpenRecordset(tblName$, dbOpenDynaset)
            If Err Then skipTab = True
            Set rsTst = Nothing
            Err.Clear
            End If
        If Not skipTab Then
            seq = seqBase_prp
            If Not skipTabPrp Then
                Err.Clear
                For Each prp In tbD.Properties
                    Select Case prp.name
                    'If (prp.Name <> "GUID") And (prp.Name <> "NameMap") Then
                        Case "GUID"
                        Case "NameMap"
                        Case "ConflictTable"
                        Case "ReplicaFilter"
                        Case Else
                                    rs.AddNew
                                    seq = seq + 1
                                    rs!TableName = tblName$
                                    rs!FieldName = "Prop nn: " & prp.name
                                    rs!FieldType = prp.Type
                                    rs!FieldPurpose = prp.value
                                    rs!FieldSeq = seq       ' reporting sequence
                                    rs.Update
                        End Select
                    Next prp
                Err.Clear
                End If
            If Not skipTabIdx Then
                seq = seqBase_idx
                pseq = seq
                For Each indx In tbD.Indexes
                    rs.AddNew
                    seq = seq + 1
                    rs!TableName = tblName$
                    rs!FieldName = "Index nn: " & indx.name
                    rs!FieldType = "Primary: " & indx.Primary & "  Unique: " & indx.Unique
                    wk$ = "Fields: "
                    For Each fld In indx.Fields
                        wk$ = wk$ & fld.name & "; "
                        Next fld
                    Err.Clear
                    rs!FieldPurpose = Left$(wk$, 250)
                    Err.Clear
                    rs!FieldSeq = seq       ' reporting sequence
                    rs.Update
                    Next indx
                    End If
            seq = seqBase_fld
            For Each fld In tbD.Fields
                rs.AddNew
                seq = seq + 1
                rs!TableName = tblName$
                rs!FieldName = fld.name
                rs!FieldType = fld.Type
                rs!FieldSize = fld.Size
                rs!FieldSeq = seq       ' reporting sequence
                rs!FieldPurpose = DBM_GetProperty(fld, "Description")
                rs!FieldAllowZ = fld.AllowZeroLength
                Err.Clear
                rs.Update
                Next fld
            End If
        Next tbD
    CommitTrans
    Set rs = Nothing
    Set tbD = Nothing
    Set DBtgt = Nothing
    wk$ = DBworkTab.name
    Set DBworkTab = Nothing
    endProc = Now

    MsgBox "Start: " & begProc & vbCrLf & "End: " & endProc & vbCrLf & "Duration (secs): " & Int(DateDiff("s", begProc, endProc))
    If Not skipReport Then
        Dim col_arg As Collection
        Set col_arg = New Collection
        col_arg.Add dTablName$      '   not including braces
        col_arg.Add wk$
        col_arg.Add tgtDb$
        glb_coll.Add col_arg       ' pass info through to Report
        Set col_arg = Nothing
        Call OpenReportAllowed("R99 DB Schema", acViewPreview)
        End If
    End Sub

Function Exclude_based_on_Name(ByVal tblName$) As Boolean
    ' look for system tables, work tables etc and block their analysis
    Exclude_based_on_Name = True
    
    If (Left$(tblName$, 4) = "MSys") Then Exit Function
    If (Left$(tblName$, 1) = "~") Then Exit Function
    ' .... put in any others here that should be suppressed
#If used_by_CD Then
    If (Left$(tblName$, 4) = "zzz ") Then Exit Function
    If (Left$(tblName$, 3) = "zz ") Then Exit Function
    If (Left$(tblName$, 1) = "_") Then Exit Function
#End If
    Exclude_based_on_Name = False
    End Function

Public Function DBM_GetProperty(tgtObj As Object, prpName As String)

    On Local Error Resume Next
    DBM_GetProperty = tgtObj.Properties(prpName)
    If Err Then
        DBM_GetProperty = ""
        End If
    End Function

Public Function DBM_SetProperty(tgtObj As Object, prpName As String, prpVal As Variant, Optional prpType = dbText) As Boolean

Const acErr_PropNotFound = 3270         ' Property was not found.
Const acErr_PropNullNotAllowed = 3385   ' User-defined properties do not support a Null value.
    
    Dim prpNew As Object

    On Local Error Resume Next
    tgtObj.Properties(prpName) = prpVal     ' Attempt to set the specified property.
    Select Case Err.Number
        Case 0
        Case acErr_PropNotFound

                    ' Create property, set its value, and append it to the Properties collection.
                    Err.Clear
                    Set prpNew = tgtObj.CreateProperty(prpName, prpType, prpVal)
                    tgtObj.Properties.Append prpNew
        Case acErr_PropNullNotAllowed
                    ' Delete  property
                    Err.Clear
                    tgtObj.Properties.Delete (prpName)
        Case Else
                    ' If different error has occurred, display message.
#If used_by_CD Then
                    System_Error_Log "DBM_SetProperty", _
                                     "PrpName: " & prpName & " PrpType: " & prpType & " PrpVal: " & prpVal
#End If
        End Select
    DBM_SetProperty = (Err = 0)
    End Function

Public Function DBM_OverrideDefault(preferredVal$, candidate$)
    
    Dim wk$
    wk = "dflt "
    If 1 = InStr(candidate$, wk) Then
        If preferredVal$ = "" Then
            preferredVal$ = Mid$(candidate$, Len(wk) + 1)
            End If
        Else
        If candidate$ <> "" Then preferredVal$ = candidate$
        End If
    DBM_OverrideDefault = preferredVal
    End Function

Public Function DBM_EditTable_FieldLookup(tgtDb$, tblName$, tgtFld$ _
                               , Optional sqlStr$ _
                               , Optional srcType$ = "dflt Table/Query" _
                               , Optional ctlType$ = "dflt Combo Box" _
                               , Optional colCount$ = "dflt 2" _
                               , Optional colWidths$ = "dflt 0" _
                               , Optional listWidth$ = "dflt 0*567twip" _
                               ) As Boolean
'   Set up a Table field as a combo box
'   can use expressions in widths colWidths$, listWidth$ ...
'   e.g.
'       6.5*567  (allows conversion of cm to twips) ... result 3685.5
'       6.5cm  (allows conversion of cm to twips) ... result 3685.5
'       0;1.5cm;6.5cm ...result 0;850.5;3685.5

    Dim fld As DAO.Field, DBtgt As DAO.Database, prp As Property
    Dim i%, j%, researchCurrent As Boolean
    Dim sql$, lSrcType$, lctlType$, lcolCount$, lcolWidths$, llistWidth$

    DBM_EditTable_FieldLookup = False
    Set DBtgt = DBM_DB_Owning_Table(tgtDb, tblName)
    Set fld = DBtgt.TableDefs(tblName).Fields(tgtFld$)
    Select Case ctlType$
        Case "Text Box":    lctlType$ = acDispControlType_TextBox    ' Text Box
        Case "List Box":    lctlType$ = acDispControlType_ListBox    ' List Box
        Case Else:          lctlType$ = acDispControlType_ComboBox    ' Combo Box
        End Select
    researchCurrent = False
    If sqlStr$ = "" Then
        sql$ = DBM_GetProperty(fld, "Rowsource")
        If sql$ = "" Then
            '   bad call ... can't not supply and not have also
            MsgBox "No sql provided, and none already defined. " & vbCrLf & _
                   "Can't edit combobox data that doesn't exist.", _
                   vbCritical + vbOKOnly, _
                   "Fatal error"
            Set fld = Nothing
            Set DBtgt = Nothing
            Exit Function
            End If
        ' get existing properties
        researchCurrent = True
        Else
        sql$ = DBM_GetProperty(fld, "Rowsource")
        If sql$ <> "" Then
            researchCurrent = True
            End If
        End If
    If researchCurrent Then
        lSrcType$ = DBM_GetProperty(fld, "RowSourceType")
        lcolCount = DBM_GetProperty(fld, "ColumnCount")
        lcolWidths$ = DBM_GetProperty(fld, "ColumnWidths")
        llistWidth$ = DBM_GetProperty(fld, "ListWidth")
        End If
    
    sql$ = DBM_OverrideDefault(sql, sqlStr$)
    lSrcType$ = DBM_OverrideDefault(lSrcType$, srcType$)
    lcolCount = Val(DBM_OverrideDefault(lcolCount, colCount))
    lcolWidths$ = DBM_OverrideDefault(lcolWidths$, colWidths$)
    Dim factor As Long, UOM$
    If lcolWidths$ = "" Then
        lcolWidths$ = "0"
        Else
        Do
            UOM$ = "cm": i = InStr(lcolWidths$, UOM$)
            If i <> 0 Then
                factor = fig_TwipsPerCm
                Else
                UOM$ = "in": i = InStr(lcolWidths$, UOM$)
                If i = 0 Then Exit Do
                factor = fig_TwipsPerInch
                End If
            j = InStrRev(Left$(lcolWidths$, i - 1), ";")
            lcolWidths$ = Left$(lcolWidths$, j) & CStr(Val(Mid$(lcolWidths$, j + 1)) * factor) & Mid$(lcolWidths$, i + Len(UOM$))
            Loop
        End If
    
    llistWidth$ = DBM_OverrideDefault(llistWidth$, listWidth$)
    UOM$ = "cm": i = InStr(llistWidth$, "cm")
    If i <> 0 Then
        factor = fig_TwipsPerCm
        Else
        UOM$ = "in": i = InStr(lcolWidths$, UOM$)
        If i <> 0 Then
            factor = fig_TwipsPerInch
            Else
            factor = 1
            UOM$ = "twips": i = InStr(llistWidth$, UOM$)
            If i = 0 Then UOM$ = "twip": i = InStr(llistWidth$, UOM$)
            End If
        End If
    If i <> 0 Then llistWidth$ = Left$(llistWidth$, i - 1)
    ' Eval in case llistWidth has an expression
    llistWidth$ = CStr(Eval(llistWidth$) * factor) & "twip"
    
    i = Val(lctlType$)
    Call DBM_SetProperty(fld, "DisplayControl", i, dbInteger)
    Call DBM_SetProperty(fld, "RowSourceType", lSrcType$, dbText)
    Call DBM_SetProperty(fld, "RowSource", sql$, dbMemo)
    Call DBM_SetProperty(fld, "BoundColumn", 1, dbInteger)
    Call DBM_SetProperty(fld, "ColumnCount", lcolCount, dbInteger)
    Call DBM_SetProperty(fld, "ColumnWidths", lcolWidths$, dbText)
    Call DBM_SetProperty(fld, "ListWidth", llistWidth$, dbText)
    Call DBM_SetProperty(fld, "LimitToList", True, dbBoolean)
    Set fld = Nothing
    Set DBtgt = Nothing
    DBM_EditTable_FieldLookup = True
    End Function

Public Function DBM_DB_Owning_Table(tgtDb As String, tblName As String) As DAO.Database
' *^* Jet-only version

    Dim wk$
    If tgtDb = "" Then
        wk = Nz(CurrentDb.TableDefs(tblName).connect, "")
        If wk = "" Then
            Set DBM_DB_Owning_Table = CurrentDb()
            Else
            wk = Mid$(wk, InStr(wk, figConnectIntro) + Len(figConnectIntro))
            Set DBM_DB_Owning_Table = OpenDatabase(wk)
            End If
        Else
        Set DBM_DB_Owning_Table = OpenDatabase(tgtDb)
        End If
    End Function

Public Function OpenReportAllowed(repName$ _
                        , Optional viewMode As AcView = acViewNormal _
                        , Optional filtName$ = "" _
                        , Optional wherecond$ = "" _
                                ) As Boolean
    
'   return True if report runs normally
'   return False if some error (e.g. No Data) is generated

    OpenReportAllowed = False
    On Error Resume Next
    DoCmd.OpenReport repName$, viewMode, filtName$, wherecond$
    If Err = 0 Then OpenReportAllowed = True
    End Function
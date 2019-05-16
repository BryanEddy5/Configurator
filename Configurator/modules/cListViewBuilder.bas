Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Const QTEMP_PREFIX As String = "qtemplistview"

'********************************************************************
'*                                                                                                   Jun 18 2007
'*    This module provides generic functionality for implementing a listview control.
'*    USAGE:  Call the load method, supplying parameters, then call the Create method.
'*
'********************************************************************
Event CreatedHeader(Header As MSComctlLib.ColumnHeader, Field As DAO.Field)
Event CreatedListItem(ListItem As MSComctlLib.ListItem, Fields As DAO.Fields)
Event CreatedListSubItem(ListSubItem As MSComctlLib.ListSubItem, Parent As MSComctlLib.ListItem, Field As DAO.Field)
Event CreatedListview(Listview As MSComctlLib.Listview)
Event CreatedSQL(ByVal Base As String, ByVal Active As String)
Event SetDateFormat(DateFormat As String)
Event SetCurrencyFormat(CurrencyFormat As String)


'********************************************************************
'*                                                                                                   Jun 18 2007
'*       Module Variables>
'*
'********************************************************************
Private WithEvents m_lvw As MSComctlLib.Listview
Attribute m_lvw.VB_VarHelpID = -1
Private WithEvents m_frm As Access.Form
Attribute m_frm.VB_VarHelpID = -1

Private m_cc As Access.CustomControl
Private m_dbs As DAO.Database
Private m_qbase As DAO.QueryDef
Private m_qactive As DAO.QueryDef

Private m_qname As String
Private m_where As String
Private m_orderby As String
Private m_dateFormat As String
Private m_currencyFormat As String
Private m_messages As String

Private m_isCreated As Boolean

'********************************************************************
'*                                                                                                   Jun 18 2007
'*       Properties>
'*
'********************************************************************
Private Property Get SQLWhere() As String
   SQLWhere = m_where
End Property
Private Property Let SQLWhere(sValue As String)
   If sValue <> m_where Then
      Set m_qactive = Nothing
      m_where = sValue
   End If
End Property

Private Property Get SQLOrderBy() As String
   SQLOrderBy = m_orderby
End Property
Private Property Let SQLOrderBy(sValue As String)
   If sValue <> m_orderby Then
      m_orderby = sValue
      Set m_qactive = Nothing
   End If
End Property

Private Property Get sql() As String
   Dim criteria As String
   'apply secondary where and orderby clauses
   If SQLWhere <> "" Then criteria = "WHERE " & SQLWhere & " "
   If SQLOrderBy <> "" Then criteria = criteria & "ORDER BY " & SQLOrderBy
   'retain the composite sql
   sql = "SELECT * FROM " & BaseQuery.name & " " & criteria
End Property

Private Property Get ColumnWidths() As Variant
   Dim hdr As MSComctlLib.ColumnHeader
   Dim tmp As String
   
   'create a list of space delimited column header widths
   For Each hdr In ListViewControl.ColumnHeaders
      tmp = tmp & " " & Round(hdr.Width, 2)
   Next
   'drop leading space
   tmp = Trim(tmp)
   ColumnWidths = Split(tmp)
   
End Property

Private Property Get AutoScaleColumns() As Boolean
   AutoScaleColumns = Not m_frm Is Nothing
End Property
Private Property Let AutoScaleColumns(bValue As Boolean)
   If Not bValue Then
      Set m_frm = Nothing
   Else
      If m_frm Is Nothing Then Set m_frm = m_cc.Parent
   End If
End Property

'********************************************************************
'*                                                                                                   Dec 15 2008
'*       Object Properties>
'*
'********************************************************************
Private Property Get Database() As DAO.Database
   If m_dbs Is Nothing Then Set m_dbs = CurrentDb
   Set Database = m_dbs
End Property

Private Property Get BaseQuery() As DAO.QueryDef
'  Exposes the base query created in the load method
   Set BaseQuery = m_qbase
End Property

Private Property Get ActiveQuery() As DAO.QueryDef
   If m_qactive Is Nothing Then
      Set m_qactive = Database.CreateQueryDef("", sql)
      RaiseEvent CreatedSQL(BaseQuery.sql, sql)
   End If
   Set ActiveQuery = m_qactive
End Property

Private Property Get ListViewControl() As MSComctlLib.Listview
   Set ListViewControl = m_lvw
End Property


'********************************************************************
'*                                                                                                   Feb 18 2007
'*       Methods>
'*
'********************************************************************
Public Function Load(ListViewControl As Access.CustomControl, sql As String, Optional ColumnWidths As String, Optional dbs As DAO.Database) As cListViewBuilder
   'save object references
   If Not dbs Is Nothing Then Set m_dbs = dbs
   
'   DeleteOldQueries Database

   Set m_cc = ListViewControl
   Set m_frm = ListViewControl.Parent
   Set m_lvw = ListViewControl.Object
   
   'enable Me to do resize handling for the parent form...
   If m_frm.OnResize <> "[Event Procedure]" Then m_frm.OnResize = "[Event Procedure]"
   
   'create the query against which where clauses and sort orders will be applied.
   m_qname = QTEMP_PREFIX & GetRandomName
   Set m_qbase = Database.CreateQueryDef(m_qname, sql)
   
   'create listview headers
   CreateHeaders GetColumns(ColumnWidths)
   'set common properties of the list view control
   SetControlProperties
   
   'assign to function
   Set Load = Me

End Function

Private Sub Initialize()
   'called by the Create() method, this routine runs once when Create() is first run
   'consumers might not have a valid WithEvents instance during the Load method
   RaiseEvent CreatedListview(m_lvw)
   m_dateFormat = "dMMMyy"
   m_currencyFormat = "Currency"
   RaiseEvent SetDateFormat(m_dateFormat)
   RaiseEvent SetCurrencyFormat(m_currencyFormat)

End Sub

Public Sub Rebuild(Optional where As Variant, Optional OrderBy As Variant)
On Error GoTo handler
   Dim fld As DAO.Field
   Dim item As MSComctlLib.ListItem
   Dim SubItem As MSComctlLib.ListSubItem
   Dim sValue As String
   
   If Not m_isCreated Then
      Initialize
      m_isCreated = True
   End If
   
   'apply these constraints if they are supplied, noting that if they are not supplied they must be removed
   If Not IsMissing(where) Then SQLWhere = where
   If Not IsMissing(OrderBy) Then SQLOrderBy = OrderBy

   'clear previous list items
   ListViewControl.ListItems.Clear
   
   'add list items
   With ActiveQuery.OpenRecordset
      'traverse recordset
      Do While Not .EOF
         'traverse fields collection
         For Each fld In .Fields
            
            'set format based on DAO.DataTypeEnum for fld.type
            Select Case fld.Type
               Case dbCurrency
                  sValue = Format(Nz(fld.value, ""), m_currencyFormat)
               Case dbDate
                  sValue = Format(Nz(fld.value, ""), m_dateFormat)
               Case Else
                  sValue = Nz(fld.value, "")
            End Select
            
            If fld.OrdinalPosition = 0 Then
               'this is the first field, always the 'bound column', and a ListItem
               Set item = m_lvw.ListItems.Add(, , sValue)
            Else
               'all other fields are ListSubItems
               Set SubItem = item.ListSubItems.Add(, fld.name, sValue)
               RaiseEvent CreatedListSubItem(SubItem, item, fld)
            End If
            
         Next fld
         RaiseEvent CreatedListItem(item, .Fields)
         .MoveNext
      Loop
      .Close
   End With
   
final:
   Exit Sub

handler:
   If Err = 3021 Or Err = 35604 Then   'No Current Record, and The first column in a ListView control must be left aligned
      'the recordset is empty, so attempts to set field alignments for $ and % fails
      Resume Next
   Else
      Err.Raise Err, Err.Source & " " & TypeName(Me) & ".Rebuild()"
      Resume final
   End If
   
End Sub

Private Sub SetControlProperties()
   'if the parameter was not supplied, attempt to use the the item passed in during Load
   'again check for the existence of the control
   If Not ListViewControl Is Nothing Then
      'set properties
      With ListViewControl
         .AllowColumnReorder = True
         .FullRowSelect = True
         .Font.name = "Tahoma"
         .Font.Size = 8
         .GridLines = True
         .HideSelection = False
         .LabelEdit = lvwManual
         .MultiSelect = False
         .OLEDragMode = ccOLEDragManual
         .OLEDropMode = ccOLEDropNone
         .Sorted = False
         .view = lvwReport
      End With
   Else
      Err.Raise -481723694, TypeName(Me) & ".SetControlProperties()", "The control 'Is Nothing'."
   End If

End Sub

'********************************************************************
'*                                                                                                   Jun 18 2007
'*       List View Event Handlers>
'*
'********************************************************************
Private Sub m_lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'  Handles a click on a column header, which has the effect of sorting the list on that column
'  Or reverse sorting the list if this is the second click on the same column header
   
   If InStr(1, SQLOrderBy, ColumnHeader.tag) = 2 Or SQLOrderBy = "" Then
      'this is a second click on the same field that was previously defined
      'in this case reverse the previous sort order
      If Right(SQLOrderBy, 4) <> "Desc" Then
         SQLOrderBy = "[" & ColumnHeader.tag & "] Desc"
      Else
         SQLOrderBy = "[" & ColumnHeader.tag & "]"
      End If
   Else
      'this is the first click on a new column, so sort ascending on the specified field
      SQLOrderBy = "[" & ColumnHeader.tag & "]"
   End If
   
   Me.Rebuild

End Sub

'********************************************************************
'*                                                                                                                                   Mar 19 2012
'*       Class Events
'*
'********************************************************************
Private Sub Class_Terminate()
On Error GoTo ErrorHandler
   'attempt to delete and temp queries created by this class
   
   With dbs.OpenRecordset( _
      "SELECT Name FROM MSysObjects " & _
      "WHERE Type = 5 AND Left(Name, 13) = 'qtemplistview'")
      
      Do While Not .EOF
         dbs.QueryDefs.Delete !name
         .MoveNext
      Loop
      
      .Close
   End With
ErrorExit:
Exit Sub
   
ErrorHandler:
   Resume Next
   
End Sub


'********************************************************************
'*                                                                                                                                   Sep 2 2010
'*       Form Resize Handler>
'*
'********************************************************************
Private Sub m_frm_Resize()
   Static lastWidth As Long
   Dim tmp As Single
   
   If lastWidth = 0 Then lastWidth = m_frm.InsideWidth
   tmp = m_frm.InsideWidth / lastWidth
   RescaleColumns tmp
   lastWidth = m_frm.InsideWidth
End Sub


'********************************************************************
'*                                                                                                   Jun 18 2007
'*       Utilities>
'*
'********************************************************************
Private Sub CreateHeaders(Cols As Variant)
   Dim Header As MSComctlLib.ColumnHeader
   Dim fld As DAO.Field
   Dim colScale As Single
   
   colScale = GetColumnScale(Cols)
   
   ListViewControl.ColumnHeaders.Clear
   
   With BaseQuery
      If .Fields.Count <> UBound(Cols) + 1 Then
         Err.Raise -41829376, TypeName(Me) & ".CreateHeaders()", "Recordset field count does not match the number of supplied column widths."
      Else
         'add column headers
         For Each fld In .Fields
            'create a header for the current field
            Set Header = m_lvw.ColumnHeaders.Add(, fld.name, fld.name)
            With Header
               'save the name of the SourceField in the tag field of the columnheader object
               'this may be used to apply sorting to the recordsource of a report
               .tag = fld.name
               'enforce column widths
               .Width = Cols(fld.OrdinalPosition) * colScale
               'left align first col
               If fld.OrdinalPosition = 0 Then
                  'default to left aligned
               Else
                  'set column alignments based on fld.type
                  Select Case fld.Type
                     Case dbDate
                        .Alignment = lvwColumnCenter
                     Case dbText, dbMemo
                        .Alignment = lvwColumnLeft
                     Case Else      ' dbCurrency, dbSingle, dbDouble, dbLong, dbByte, dbInteger
                        .Alignment = lvwColumnRight
                  End Select
               End If
            End With
            RaiseEvent CreatedHeader(Header, fld)
         Next fld
      End If
      .Close
   End With
   
End Sub

Friend Function GetColumns(Widths As String)
   Dim var
   Dim i As Integer
   
   var = Split(Widths, ";")
   If UBound(var) + 1 = m_qbase.Fields.Count Then
      'no problem
   ElseIf UBound(var) + 1 < m_qbase.Fields.Count Then
      'the user supplied column widths are fewer than the fields that exist in the query
      'so add column widths with a length of 1
      For i = UBound(var) + 1 To m_qbase.Fields.Count - 1
         ReDim Preserve var(UBound(var) + 1)
         If UBound(var) = 0 Then
            'this is the first field so default to zero width
            var(0) = 0
         Else
            var(UBound(var)) = 1
         End If
      Next
   ElseIf UBound(var) + 1 > m_qbase.Fields.Count Then
      ReDim Preserve var(m_qbase.Fields.Count - 1)
   End If
   
   'ensure that all columns contain numeric data
   For i = LBound(var) To UBound(var)
      If Not IsNumeric(var(i)) Or StringContainsCharacters(CStr(var(i)), ",-|~/?<>'""") Then
         Err.Raise 5, , _
            "Invalid ColumnWidths: """ & Widths & """" & vbCrLf & _
            "Please specify semi-colon delimited proportional column widths." & vbCrLf & _
            "   eg: ""0;4;2;1""  " & vbCrLf & _
            "If this parameter is omitted then ""0;1;1;1;1;..."" is assumed."
      End If
   Next
   
   GetColumns = var
   
   Exit Function

handler:
   Err.Raise Err, "cListViewBuilder"

End Function

Private Function GetColumnScale(Cols As Variant) As Single
   Dim var
   Dim sum As Single
   
   For Each var In Cols
      'sum the column widths, which might be proportional
      sum = sum + var
   Next
   If sum > 0 Then GetColumnScale = (m_cc.Width - 0.3 * 1440) / sum
   
End Function

Private Sub RescaleColumns(scaleTo As Single)
   Dim Col As MSComctlLib.ColumnHeader
   For Each Col In m_lvw.ColumnHeaders
      Col.Width = Col.Width * scaleTo
   Next
End Sub

Private Function GetRandomName() As String
   Dim tmp As String
   tmp = LibSys.FSO.GetTempName & LibSys.FSO.GetTempName
   tmp = Replace(tmp, "rad", "")
   tmp = Replace(tmp, ".tmp", "")
   GetRandomName = tmp
End Function

Private Function StringContainsCharacters(Text As String, list As String) As Boolean
   Dim var
   Dim tmp() As Byte
   tmp = list
   
   For Each var In tmp
      If InStr(Text, Chr(var)) Then
         StringContainsCharacters = True
         Exit Function
      End If
   Next
   
End Function

Private Sub DeleteOldQueries(dbs As DAO.Database)
   Dim rst As DAO.Recordset
   
   Set rst = dbs.OpenRecordset( _
      "SELECT Name FROM MSysObjects " & _
      "WHERE Left(Name, " & Len(QTEMP_PREFIX) & ") = '" & QTEMP_PREFIX & "' " & _
         "AND Type = 5")
   With rst
      Do While Not .EOF
         dbs.QueryDefs.Delete !name
         .MoveNext
      Loop
      .Close
   End With

End Sub
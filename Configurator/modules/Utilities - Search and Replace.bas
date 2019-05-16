Option Compare Database
Option Explicit

'   Family of routines to assist with change of field and table names
'   in a mature Access db.

'   string_in_xxx_find
'   Provides facilities to scan through forms, reports and queries
'   and display to the screen (or a file) any 'hits' on name matches
'   example:
'       string_in_forms_find  "Customer Table"

'   string_in_xxx_alter
'   Provides facilities to scan through forms, reports and queries
'   and replace one string (e.g. table name) with another
'   optionally, a list of alterations can be logged
'   example:
'       string_in_reports_alter "Customer Table","Customer Table Contacts","tblCustomers"

'   log file is called <fig_logFileName> - see below - and is in the same folder
'   as the database being investigated / altered

Public Const fig_NoLogFile = 0
Public Const fig_logFileName = "change log.txt"
Public Const max_Form_Sects = 5
Public Const max_Report_Sects = 11
'
Const figTempFile_CoreName = "_delete this tempfile_ "  ' used by TempQuery (Name) build/delete

Public Const err_ItemNotInThisCollection = 3265

'================== the API calls are merely to try for pretty-looking
'================== change confirmation MsgBox messages
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90

Type Size
    cx As Long
    cy As Long
    End Type

Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, _
                                             ByVal nIndex As Long) As Long
Declare Function GetFocus Lib "user32" () As Long
Declare Function GetTextExtentPoint Lib "gdi32" Alias _
                "GetTextExtentPointA" (ByVal hDC As Long, ByVal lpsz As String, _
                ByVal cbString As Long, lpSIZE As Size) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
                                            ByVal hDC As Long) As Long
'
'================ end of API stuff

Public Sub String_In_Forms_Find(ByVal searchFor$ _
                     , Optional ByVal findThemAll As Boolean = True _
                     , Optional ByVal thisProperty As String = "" _
                     , Optional ByVal soloObjName$ = "" _
                     , Optional ByVal logToFile As Boolean = False _
                                )
'*******************************************
'Purpose:   Find and list the names of forms that contain at least one instance of 'searchFor$'
'           if all instances of 'searchFor$' in a form are required to be listed, set findThemAll = True
'           if only a particular Property is to be searched, set the name in thisProperty
'           if only a particular form is to be searched, set the name in soloObjName$
'           if results to be written to file, set logToFile = True
'
'Author:    CD
'Date:      Nov 1997, rev Aug 2000
'Called by: Free-standing utility, used in dev. phase to locate name(obj,method) usage, etc
'           see related String_In_Forms_Alter, String_In_Reports_Find
'*******************************************

    Dim obj As AccessObject, dbs As Object
    Dim frm As Form, ctl As control, prp As Property, sect As Section
    Dim Found As Boolean, hit As Long, sought As Long
    Dim doIt As Boolean
    
    Dim wk$, i%, prpVal$, pth$, chn%

    hit = 0: sought = 0: chn = fig_NoLogFile
    If logToFile Then
        chn = FreeFile
        pth = CurrentDb.name
        i = InStrRev(pth, "\"): pth = Left$(pth, i) & fig_logFileName
        Open pth$ For Append As #chn
        Print #chn, "String_In_Forms_Find - "; Now; "   "; CurrentDb.name
        Print #chn, "                        searchFor$: "; searchFor$
        Print #chn, "                       findThemAll: "; findThemAll
        Print #chn, "                       thisProperty: "; thisProperty$
        Print #chn, "                   soloObjName$: "; soloObjName$
        End If
    Set dbs = Application.CurrentProject
    On Local Error Resume Next
    For Each obj In dbs.AllForms
        If (soloObjName$ = "") Or (soloObjName$ = obj.name) Then
            sought = sought + 1
            If obj.IsLoaded Then DoCmd.Close acForm, obj.name, acSaveYes
            Err.Clear
            Do
                DoCmd.OpenForm obj.name, acDesign, , , , acHidden
                If Err = 0 Then Exit Do
                ' probably subform of open form
                If Forms.Count = 0 Then
                    MsgBox "Error on document Open action" & _
                           "  Err: " & Err.Number & " - " & Err.Description _
                           , , "Can't open document"
                    Exit For
                    End If
                DoCmd.Close acForm, Forms(0).name
                Err.Clear
                Loop
            Set frm = Forms(obj.name)
            Found = False
            ' first scan each Form property
            ' then, for each Section in the Form, scan each Section property
            ' then, for each Control in the Form, scan each Control property
            For Each prp In frm.Properties
                Err.Clear
                Select Case prp.name
                    Case "LogicalPageWidth" ' skip to avoid 'printer missing'-type messages
                    Case Else
                            prpVal$ = Nz(prp.value, "")
                            If Err = 0 Then
                                If thisProperty = "" Or thisProperty = prp.name Then
                                    If InStr(prpVal$, searchFor) Then
                                        wk$ = "     " & "FORM" & " - " & prp.name & ": " & prpVal$
                                        Feedback Found, chn, frm.name, wk$
                                        If Not findThemAll Then Exit For
                                        End If
                                    End If
                                End If
                    End Select
                Next prp

            If findThemAll Or Not Found Then
'               Traverse Sections "collection" of each form.
                For i = 0 To max_Form_Sects - 1
                    Err.Clear
                    Set sect = frm.Section(i)
                    If Err = 0 Then
                        For Each prp In sect.Properties
                            Err.Clear
                            prpVal$ = Nz(prp.value, "")
                            If Err = 0 Then
                                If thisProperty = "" Or thisProperty = prp.name Then
                                    If InStr(prpVal$, searchFor) Then
                                        wk$ = "     " & "SECTION " & CStr(i) & " - " & prp.name & ": " & prpVal$
                                        Feedback Found, chn, frm.name, wk$
                                        If Not findThemAll Then Exit For
                                        End If
                                    End If
                                End If
                            Next prp
                        End If
                    Set sect = Nothing
                    Next i
                End If
            If findThemAll Or Not Found Then
'               Traverse Controls collection of each form.
                For Each ctl In frm.Controls
                    For Each prp In ctl.Properties
                        Err.Clear
                        prpVal$ = Nz(prp.value, "")
                        If Err = 0 Then
                            If thisProperty = "" Or thisProperty = prp.name Then
                                If InStr(prpVal$, searchFor) Then
                                    wk$ = "     " & ctl.name & " - " & prp.name & ": " & prpVal$
                                    Feedback Found, chn, frm.name, wk$
                                    If Not findThemAll Then Exit For
                                    End If
                                End If
                            End If
                        Next prp
                    If Found And Not findThemAll Then Exit For
                    Next ctl
                End If
            If Found Then
                hit = hit + 1
                End If
            DoCmd.Close acForm, obj.name, acSaveNo
            End If
        Next obj
    Set dbs = Nothing
    wk$ = "Examined: " & CStr(sought) & "  Found: " & CStr(hit)
    If chn <> fig_NoLogFile Then
        Feedback True, chn, "", wk$
        Close #chn
        wk$ = wk$ & "  Written to: " & pth$
        End If
    Debug.Print wk$
    End Sub

Public Sub String_In_Forms_Alter(ByVal searchFor$ _
                               , ByVal replaceWith$ _
                      , Optional ByVal exceptFor$ = "" _
                      , Optional ByVal soloObjName$ = "" _
                      , Optional ByVal logToFile As Boolean = False _
                      , Optional ByVal Confirm As Boolean = False _
                      )
'*******************************************
'Purpose:   Find occurrences of 'searchFor$' and replace them with 'replaceWith$'
'           if 'searchFor$' occurs as part of 'exceptFor$', then skip the replacement
'           if only a particular form is to be searched, set the name in soloObjName$
'           if results to be written to file, set logToFile = True
'           if each hit requires confirmation, set confirm = True
'
'Author:    CD
'Date:      Nov 1997, rev Aug 2000
'Called by: Free-standing utility, used in dev. phase to locate name(obj,method) usage, etc
'           see related String_In_Forms_Find, String_In_Reports_Find
'*******************************************

    Dim obj As AccessObject, dbs As Object
    Dim frm As Form, ctl As control, prp As Property, sect As Section
    Dim Found As Boolean, hit As Long, sought As Long
    Dim doIt As Boolean, abort As Boolean
    Dim mtchPt%, excePt%, stpt%
    Dim wk$, i%, prpVal$, pth$, chn%

    hit = 0: sought = 0: chn = fig_NoLogFile: abort = False
    If logToFile Then
        chn = FreeFile
        pth = CurrentDb.name
        i = InStrRev(pth, "\"): pth = Left$(pth, i) & fig_logFileName
        Open pth$ For Append As #chn
        Print #chn, "String_In_Forms_Alter - "; Now; "   "; CurrentDb.name
        Print #chn, "                        searchFor$: "; searchFor$
        Print #chn, "                        exceptFor$: "; exceptFor$
        Print #chn, "                      replaceWith$: "; replaceWith$
        Print #chn, "                    soloObjName$: "; soloObjName$
        End If
    Set dbs = Application.CurrentProject
    On Local Error Resume Next
    For Each obj In dbs.AllForms
        sought = sought + 1
        If soloObjName$ = "" Or soloObjName$ = obj.name Then
            If obj.IsLoaded Then DoCmd.Close acForm, obj.name, acSaveYes
            Err.Clear
            Do
                DoCmd.OpenForm obj.name, acDesign, , , , acHidden
                If Err = 0 Then Exit Do
                ' probably subform of open form
                If Forms.Count = 0 Then
                    MsgBox "Error on document Open action" & _
                           "  Err: " & Err.Number & " - " & Err.Description _
                           , , "Can't open document"
                    Exit For
                    End If
                DoCmd.Close acForm, Forms(0).name
                Err.Clear
                Loop
            Set frm = Forms(obj.name)
            Found = False
            For Each prp In frm.Properties
                Err.Clear
                prpVal$ = Nz(prp.value, "")
                If Err = 0 Then
                    stpt = 0
                    Do
                        mtchPt = InStr(stpt + 1, prpVal$, searchFor)
                        If mtchPt = 0 Then Exit Do
                        doIt = True
                        If exceptFor$ <> "" Then
                            excePt = InStr(stpt + 1, prpVal$, exceptFor$)
                            doIt = (mtchPt <> excePt)
                            If Not doIt Then stpt = mtchPt + Len(exceptFor$) - 1
                            End If
                        If doIt Then
                            If Confirm Then
                                wk$ = "      Form: " & frm.name & vbCrLf _
                                    & "Property: " & prp.name
                                doIt = ConfirmedChange(wk$, prpVal$, mtchPt, searchFor, abort, Confirm)
                                If abort Then Exit For
                                If Not doIt Then stpt = mtchPt + 1
                                Else
                                doIt = True
                                End If
                            If doIt Then
                                wk$ = "     " & "FORM" & " - " & prp.name & ": " & prpVal$
                                Feedback Found, chn, frm.name, wk$
                                prpVal$ = Left$(prpVal$, mtchPt - 1) & replaceWith$ & Mid$(prpVal$, mtchPt + Len(searchFor$))
                                prp.value = prpVal$
                                stpt = mtchPt + Len(replaceWith$) - 1
                                End If
                            End If
                        Loop
                    End If
                Next prp
            If abort Then Exit For
            For i = 0 To max_Form_Sects - 1
                Err.Clear
                Set sect = frm.Section(i)
                If Err = 0 Then
                    For Each prp In sect.Properties
                        Err.Clear
                        prpVal$ = Nz(prp.value, "")
                        If Err = 0 Then
                            stpt = 0
                            Do
                                mtchPt = InStr(stpt + 1, prpVal$, searchFor)
                                If mtchPt = 0 Then Exit Do
                                doIt = True
                                If exceptFor$ <> "" Then
                                    excePt = InStr(stpt + 1, prpVal$, exceptFor$)
                                    doIt = (mtchPt <> excePt)
                                    If Not doIt Then stpt = mtchPt + Len(exceptFor$) - 1
                                    End If
                                If doIt Then
                                    If Confirm Then
                                        wk$ = "      Form: " & frm.name & vbCrLf _
                                            & "  Section: " & CStr(i) & vbCrLf _
                                            & "Property: " & prp.name
                                        doIt = ConfirmedChange(wk$, prpVal$, mtchPt, searchFor, abort, Confirm)
                                        If abort Then Exit For
                                        If Not doIt Then stpt = mtchPt + 1
                                        Else
                                        doIt = True
                                        End If
                                    If doIt Then
                                        wk$ = "     " & "SECTION " & CStr(i) & " - " & prp.name & ": " & prpVal$
                                        Feedback Found, chn, frm.name, wk$
                                        prpVal$ = Left$(prpVal$, mtchPt - 1) & replaceWith$ & Mid$(prpVal$, mtchPt + Len(searchFor$))
                                        prp.value = prpVal$
                                        stpt = mtchPt + Len(replaceWith$) - 1
                                        End If
                                    End If
                                Loop
                            End If
                        Next prp
                    If abort Then Exit For
                    End If
                Set sect = Nothing
                Next i
            If abort Then Exit For
            For Each ctl In frm.Controls
                For Each prp In ctl.Properties
                    Err.Clear
                    prpVal$ = Nz(prp.value, "")
                    If Err = 0 Then
                        stpt = 0
                        Do
                            mtchPt = InStr(stpt + 1, prpVal$, searchFor)
                            If mtchPt = 0 Then Exit Do
                            doIt = True
                            If exceptFor$ <> "" Then
                                excePt = InStr(stpt + 1, prpVal$, exceptFor$)
                                doIt = (mtchPt <> excePt)
                                If Not doIt Then stpt = mtchPt + Len(exceptFor$) - 1
                                End If
                            If doIt Then
                                If Confirm Then
                                    wk$ = "      Form: " & frm.name & vbCrLf _
                                        & "  Control: " & ctl.name & vbCrLf _
                                        & "Property: " & prp.name
                                    doIt = ConfirmedChange(wk$, prpVal$, mtchPt, searchFor, abort, Confirm)
                                    If abort Then Exit For
                                    If Not doIt Then stpt = mtchPt + 1
                                    Else
                                    doIt = True
                                    End If
                                If doIt Then
                                    wk$ = "     " & ctl.name & " - " & prp.name & ": " & prpVal$
                                    Feedback Found, chn, frm.name, wk$
                                    prpVal$ = Left$(prpVal$, mtchPt - 1) & replaceWith$ & Mid$(prpVal$, mtchPt + Len(searchFor$))
                                    prp.value = prpVal$
                                    stpt = mtchPt + Len(replaceWith$) - 1
                                    End If
                                End If
                            Loop
                        End If
                    Next prp
                If abort Then Exit For
                Next ctl
            If abort Then Exit For
            wk$ = obj.name
            If Found Then
                hit = hit + 1
                DoCmd.Close acForm, wk$, acSaveYes
                Else
                DoCmd.Close acForm, wk$
                End If
            End If
        Next obj
    Set dbs = Nothing
    wk$ = "Examined: " & CStr(sought) & "  Found: " & CStr(hit)
    If abort Then wk$ = wk$ & "  (** aborted **)"
    If chn <> fig_NoLogFile Then
        Feedback True, chn, "", wk$
        Close #chn
        wk$ = wk$ & "  Written to: " & pth$
        End If
    Debug.Print wk$
    End Sub

Public Sub String_In_Queries_Alter(searchFor$, replaceWith$, Optional exceptFor$ = "")

'   replace any searchFor$ in query.SQL with replaceWith$
'   unless exceptFor$ (a superset of searchFor$) is specified and matches
'          coincident with this match of searchFor$

'   before amending the query ...
'   the original form of the query is written as "_was "+queryname
'   consequently
'   this routine does not examine any queries the name of which starts with "_was "

    Dim rs As DAO.Recordset, prm As Parameter, qryD As DAO.QueryDef, nqryD As DAO.QueryDef
    Dim sql$
    sql$ = ""
    sql$ = sql$ & " PARAMETERS [searchFor] Text;"
    sql$ = sql$ & " SELECT [searchFor] AS [String], " & _
                          "MSysObjects.Name AS QueryName, " & _
                          "InStr([Expression],[searchFor]) AS [In Expression], " & _
                          "InStr([Name1],[searchFor]) AS [In Name1], " & _
                          "InStr([Name2],[searchFor]) AS [In Name2]"
    sql$ = sql$ & " FROM MSysQueries INNER JOIN MSysObjects ON MSysQueries.ObjectId = MSysObjects.Id"
    sql$ = sql$ & " WHERE (((InStr([Expression],[searchFor]))<>0)) " & _
                      "OR (((InStr([Name1],[searchFor]))<>0)) " & _
                      "OR (((InStr([Name2],[searchFor]))<>0));"
    
    Set qryD = CurrentDb.CreateQueryDef("", sql$)
    Set prm = qryD.Parameters!searchFor
    prm = searchFor$
    Set rs = qryD.OpenRecordset(dbOpenForwardOnly)
    Set prm = Nothing
    Set qryD = Nothing
    Do
        If rs.EOF Then Exit Do
        If 1 = InStr(rs!QueryName, "_was ") Then
            Else
            Debug.Print rs!QueryName
            Set qryD = CurrentDb.QueryDefs(rs!QueryName)
            Set nqryD = CurrentDb.CreateQueryDef("_was " & rs!QueryName, qryD.sql)
            Set nqryD = Nothing
            qryD.sql = Altered_Val$(qryD.sql, searchFor$, replaceWith$, exceptFor$)
            End If
        rs.MoveNext
        Loop
    End Sub

Public Function Altered_Val$(tgt$, searchFor$, replaceWith$, exceptFor$)

'   replace any searchFor$ in tgt$ with replaceWith$
'   unless exceptFor$ (a superset of searchFor$) is specified and matches
'          coincident with this match of searchFor$
'   e.g. searchFor$     Smith
'        exceptFor$     Smithson
'        replaceWith$   Jones
'   ==> Convert any Smith to Jones except for Smithson
'   e.g. searchFor$     Customer
'        exceptFor$     Customer Contact
'        replaceWith$   tblCustomer
'   ==> Convert any Customer to tblCustomer except for Customer Contact

    Dim wk$, wka$, i As Integer, j As Integer
    wka$ = "": wk = tgt$
    Do
        i = InStr(wk, searchFor$)
        If i = 0 Then
            wka$ = wka$ & wk
            Exit Do
            End If
        j = 0
        If exceptFor$ <> "" Then
            j = InStr(wk, exceptFor$)
            If j <> i Then j = 0
            End If
        If j <> 0 Then
            wka$ = wka$ & Left$(wk, i - 1) & exceptFor$
            wk$ = Mid$(wk$, i + Len(exceptFor$))
            Else
            wka$ = wka$ & Left$(wk, i - 1) & replaceWith$
            wk$ = Mid$(wk$, i + Len(searchFor$))
            End If
        Loop
    Altered_Val$ = wka
    End Function

Public Sub String_In_Queries_Find(searchFor$)

'   find and list the names of Queries that contain at least one instance of 'searchFor$'

'   creates as output a Query that has one row for each query that
'   contains the searchFor string
'       col 1:      searchFor value
'       col 2:      query name of qualifying query
'       col 3:      character offset of where searchFor$ occurs in Expression (query SQL)
'       col 4:      character offset of where searchFor$ occurs Name1
'       col 5:      character offset of where searchFor$ occurs Name2
    
    Dim sql$, qryName$, wk$, i%, j%, srch$, repl$
    sql$ = ""
    sql$ = sql$ & " SELECT [searchFor] AS [String], " & _
                          "MSysObjects.Name AS QueryName, " & _
                          "InStr([Expression],[searchFor]) AS [In Expression], " & _
                          "InStr([Name1],[searchFor]) AS [In Name1], " & _
                          "InStr([Name2],[searchFor]) AS [In Name2]"
    sql$ = sql$ & " FROM MSysQueries INNER JOIN MSysObjects ON MSysQueries.ObjectId = MSysObjects.Id"
    sql$ = sql$ & " WHERE (((InStr([Expression],[searchFor]))<>0)) " & _
                      "OR (((InStr([Name1],[searchFor]))<>0)) " & _
                      "OR (((InStr([Name2],[searchFor]))<>0));"
    i = 0: j = 1
    srch$ = "[searchFor]"
    repl$ = SQL_QuoteProtect(searchFor$)
    Do
        i = InStr(j, sql, srch$)
        If i = 0 Then
            wk = wk & Mid$(sql$, j)
            Exit Do
            End If
        wk = wk & Mid$(sql$, j, i - j) & repl$
        j = i + Len(srch$)
        Loop
    sql$ = wk
    qryName$ = TempQuery_NewName("")
    If Made_New_Query(qryName$, sql$) Then
        DoCmd.OpenQuery qryName$
        End If
    End Sub

Public Sub String_In_Reports_Find(ByVal searchFor$ _
                       , Optional ByVal findThemAll As Boolean = True _
                       , Optional ByVal thisProperty As String = "" _
                       , Optional ByVal soloObjName$ = "" _
                       , Optional ByVal logToFile As Boolean = False _
                                )
'*******************************************
'Purpose:   Find and list the names of Reports that contain at least one instance of 'searchFor$'
'           if all instances of 'searchFor$' in a Report are required to be listed, set findThemAll = True
'           if only a particular Property is to be searched, set the name in thisProperty
'           if only a particular Report is to be searched, set the name in soloObjName$
'           if results to be written to file, set logToFile = True
'
'Author:    CD
'Date:      Nov 1997, rev Aug 2000
'Called by: Free-standing utility, used in dev. phase to locate name(obj,method) usage, etc
'           see related String_In_Reports_Alter, String_In_Reports_Find
'*******************************************

    Dim obj As AccessObject, dbs As Object
    Dim Rep As Report, ctl As control, prp As Property, sect As Section
    Dim Found As Boolean, hit As Long, sought As Long
    Dim doIt As Boolean
    
    Dim wk$, i%, prpVal$, pth$, chn%

    hit = 0: sought = 0: chn = fig_NoLogFile
    If logToFile Then
        chn = FreeFile
        pth = CurrentDb.name
        i = InStrRev(pth, "\"): pth = Left$(pth, i) & fig_logFileName
        Open pth$ For Append As #chn
        Print #chn, "String_In_Reports_Alter - "; Now; "   "; CurrentDb.name
        Print #chn, "                          searchFor$: "; searchFor$
        Print #chn, "                         findThemAll: "; findThemAll
        Print #chn, "                        thisProperty: "; thisProperty$
        Print #chn, "                        soloObjName$: "; soloObjName$
        End If
    Application.Echo False      ' try to keep screen noise down (not entirely successful)
    Set dbs = Application.CurrentProject
    On Local Error Resume Next
    For Each obj In dbs.AllReports
        If (soloObjName$ = "") Or (soloObjName$ = obj.name) Then
            sought = sought + 1
            If obj.IsLoaded Then DoCmd.Close acReport, obj.name, acSaveYes
            Err.Clear
            Do
                DoCmd.OpenReport obj.name, acViewDesign
                If Err = 0 Then Exit Do
                ' probably subReport of open Report
                If Reports.Count = 0 Then
                    MsgBox "Error on document Open action" & _
                           "  Err: " & Err.Number & " - " & Err.Description _
                           , , "Can't open document"
                    Exit For
                    End If
                DoCmd.Close acReport, Reports(0).name
                Err.Clear
                Loop
            Set Rep = Reports(obj.name)
            Found = False
            ' first scan each Report-level property
            ' then scan each Section-level property
            ' then scan each Control-level property
            For Each prp In Rep.Properties
                Err.Clear
                Select Case prp.name
                    Case "LogicalPageWidth" ' skip to avoid 'printer missing'-type messages
                    Case Else
                            prpVal$ = Nz(prp.value, "")
                            If Err = 0 Then
                                If thisProperty = "" Or thisProperty = prp.name Then
                                    If InStr(prpVal$, searchFor) Then
                                        wk$ = "     " & "REPORT" & " - " & prp.name & ": " & prpVal$
                                        Feedback Found, chn, Rep.name, wk$
                                        If Not findThemAll Then Exit For
                                        End If
                                    End If
                                End If
                    End Select
                Next prp
                
            If findThemAll Or Not Found Then
'               Traverse Sections "collection" of each Report.
                For i = 0 To max_Report_Sects - 1
                    Err.Clear
                    Set sect = Rep.Section(i)
                    If Err = 0 Then
                        For Each prp In sect.Properties
                            Err.Clear
                            prpVal$ = Nz(prp.value, "")
                            If Err = 0 Then
                                If thisProperty = "" Or thisProperty = prp.name Then
                                    If InStr(prpVal$, searchFor) Then
                                        wk$ = "     " & "Section " & CStr(i) & " - " & prp.name & ": " & prpVal$
                                        Feedback Found, chn, Rep.name, wk$
                                        If Not findThemAll Then Exit For
                                        End If
                                    End If
                                End If
                            Next prp
                        End If
                    Set sect = Nothing
                    Next i
                End If
            If findThemAll Or Not Found Then
'               Traverse Controls collection of each Report.
                For Each ctl In Rep.Controls
                    For Each prp In ctl.Properties
                        Err.Clear
                        prpVal$ = Nz(prp.value, "")
                        If Err = 0 Then
                            If thisProperty = "" Or thisProperty = prp.name Then
                                If InStr(prpVal$, searchFor) Then
                                    wk$ = "     " & ctl.name & " - " & prp.name & ": " & prpVal$
                                    Feedback Found, chn, Rep.name, wk$
                                    If Not findThemAll Then Exit For
                                    End If
                                End If
                            End If
                        Next prp
                    If Found And Not findThemAll Then Exit For
                    Next ctl
                End If
            If Found Then
                hit = hit + 1
                End If
            DoCmd.Close acReport, obj.name
            End If
        Next obj
    Set dbs = Nothing
    Application.Echo True
    wk$ = "Examined: " & CStr(sought) & "  Found: " & CStr(hit)
    If chn <> fig_NoLogFile Then
        Feedback True, chn, "", wk$
        Close #chn
        wk$ = wk$ & "  Written to: " & pth$
        End If
    Debug.Print wk$
    End Sub

Public Sub String_In_Reports_Alter(ByVal searchFor$ _
                                 , ByVal replaceWith$ _
                        , Optional ByVal exceptFor$ = "" _
                        , Optional ByVal soloObjName$ = "" _
                        , Optional ByVal logToFile As Boolean = False _
                        , Optional ByVal Confirm As Boolean = False _
                        )
'*******************************************
'Purpose:   Find occurrences of 'searchFor$' and replace them with 'replaceWith$'
'           if 'searchFor$' occurs as part of 'exceptFor$', then skip the replacement
'           if only a particular Report is to be searched, set the name in soloObjName$
'           if results to be written to file, set logToFile = True
'           if each hit requires confirmation, set confirm = True
'
'Author:    CD
'Date:      Nov 1997, rev Aug 2000
'Called by: Free-standing utility, used in dev. phase to locate name(obj,method) usage, etc
'           see related String_In_Forms_Alter, String_In_Reports_Find
'*******************************************

    Dim obj As AccessObject, dbs As Object
    Dim Rep As Report, ctl As control, prp As Property, sect As Section
    Dim Found As Boolean, hit As Long, sought As Long
    Dim doIt As Boolean, abort As Boolean
    Dim mtchPt%, excePt%, stpt%
    Dim wk$, i%, prpVal$, pth$, chn%

    Set dbs = Application.CurrentProject
    hit = 0: sought = 0: chn = fig_NoLogFile: abort = False
    If logToFile Then
        chn = FreeFile
        pth = CurrentDb.name
        i = InStrRev(pth, "\"): pth = Left$(pth, i) & fig_logFileName
        Open pth$ For Append As #chn
        Print #chn, "String_In_Reports_Alter - "; Now; "   "; CurrentDb.name
        Print #chn, "                          searchFor$: "; searchFor$
        Print #chn, "                          exceptFor$: "; exceptFor$
        Print #chn, "                        replaceWith$: "; replaceWith$
        Print #chn, "                        soloObjName$: "; soloObjName$
        End If
    Application.Echo False      ' try to keep screen noise down (not entirely successful)
    On Local Error Resume Next
    For Each obj In dbs.AllReports
        sought = sought + 1
        If soloObjName$ = "" Or soloObjName$ = obj.name Then
            If obj.IsLoaded Then DoCmd.Close acReport, obj.name, acSaveYes
            Err.Clear
            Do
                DoCmd.OpenReport obj.name, acViewDesign
                If Err = 0 Then Exit Do
                ' probably subReport of open Report
                If Reports.Count = 0 Then
                    MsgBox "Error on document Open action" & _
                           "  Err: " & Err.Number & " - " & Err.Description _
                           , , "Can't open document"
                    Exit For
                    End If
                DoCmd.Close acReport, Reports(0).name
                Err.Clear
                Loop
            Set Rep = Reports(obj.name)
            Found = False
            For Each prp In Rep.Properties
                Err.Clear
                prpVal$ = Nz(prp.value, "")
                If Err = 0 Then
                    stpt = 0
                    Do
                        mtchPt = InStr(stpt + 1, prpVal$, searchFor)
                        If mtchPt = 0 Then Exit Do
                        doIt = True
                        If exceptFor$ <> "" Then
                            excePt = InStr(stpt + 1, prpVal$, exceptFor$)
                            doIt = (mtchPt <> excePt)
                            If Not doIt Then stpt = mtchPt + Len(exceptFor$) - 1
                            End If
                        If doIt Then
                            If Confirm Then
                                wk$ = "   Report: " & Rep.name & vbCrLf _
                                    & "Property: " & prp.name
                                doIt = ConfirmedChange(wk$, prpVal$, mtchPt, searchFor, abort, Confirm)
                                If abort Then Exit For
                                If Not doIt Then stpt = mtchPt + 1
                                Else
                                doIt = True
                                End If
                            If doIt Then
                                wk$ = "     " & "REPORT" & " - " & prp.name & ": " & prpVal$
                                Feedback Found, chn, Rep.name, wk$
                                prpVal$ = Left$(prpVal$, mtchPt - 1) & replaceWith$ & Mid$(prpVal$, mtchPt + Len(searchFor$))
                                prp.value = prpVal$
                                stpt = mtchPt + Len(replaceWith$) - 1
                                End If
                            End If
                        Loop
                    End If
                Next prp
            If abort Then Exit For
'           Traverse Sections "collection" of each Report.
            For i = 0 To max_Report_Sects - 1
                Err.Clear
                Set sect = Rep.Section(i)
                If Err = 0 Then
                    For Each prp In sect.Properties
                        Err.Clear
                        prpVal$ = Nz(prp.value, "")
                        If Err = 0 Then
                            stpt = 0
                            Do
                                mtchPt = InStr(stpt + 1, prpVal$, searchFor)
                                If mtchPt = 0 Then Exit Do
                                doIt = True
                                If exceptFor$ <> "" Then
                                    excePt = InStr(stpt + 1, prpVal$, exceptFor$)
                                    doIt = (mtchPt <> excePt)
                                    If Not doIt Then stpt = mtchPt + Len(exceptFor$) - 1
                                    End If
                                If doIt Then
                                    If Confirm Then
                                        wk$ = "    Report: " & Rep.name & vbCrLf _
                                            & "  Section: " & CStr(i) & vbCrLf _
                                            & "Property: " & prp.name
                                        doIt = ConfirmedChange(wk$, prpVal$, mtchPt, searchFor, abort, Confirm)
                                        If abort Then Exit For
                                        If Not doIt Then stpt = mtchPt + 1
                                        Else
                                        doIt = True
                                        End If
                                    If doIt Then
                                        wk$ = "     " & "SECTION " & CStr(i) & " - " & prp.name & ": " & prpVal$
                                        Feedback Found, chn, Rep.name, wk$
                                        prpVal$ = Left$(prpVal$, mtchPt - 1) & replaceWith$ & Mid$(prpVal$, mtchPt + Len(searchFor$))
                                        prp.value = prpVal$
                                        stpt = mtchPt + Len(replaceWith$) - 1
                                        End If
                                    End If
                                Loop
                            End If
                        Next prp
                    If abort Then Exit For
                    End If
                Set sect = Nothing
                Next i
            If abort Then Exit For
            For Each ctl In Rep.Controls
                For Each prp In ctl.Properties
                    Err.Clear
                    prpVal$ = Nz(prp.value, "")
                    If Err = 0 Then
                        stpt = 0
                        Do
                            mtchPt = InStr(stpt + 1, prpVal$, searchFor)
                            If mtchPt = 0 Then Exit Do
                            doIt = True
                            If exceptFor$ <> "" Then
                                excePt = InStr(stpt + 1, prpVal$, exceptFor$)
                                doIt = (mtchPt <> excePt)
                                If Not doIt Then stpt = mtchPt + Len(exceptFor$) - 1
                                End If
                            If doIt Then
                                If Confirm Then
                                    wk$ = "   Report: " & Rep.name & vbCrLf _
                                        & "  Control: " & ctl.name & vbCrLf _
                                        & "Property: " & prp.name
                                    doIt = ConfirmedChange(wk$, prpVal$, mtchPt, searchFor, abort, Confirm)
                                    If abort Then Exit For
                                    If Not doIt Then stpt = mtchPt + 1
                                    Else
                                    doIt = True
                                    End If
                                If doIt Then
                                    wk$ = "     " & ctl.name & " - " & prp.name & ": " & prpVal$
                                    Feedback Found, chn, Rep.name, wk$
                                    prpVal$ = Left$(prpVal$, mtchPt - 1) & replaceWith$ & Mid$(prpVal$, mtchPt + Len(searchFor$))
                                    prp.value = prpVal$
                                    stpt = mtchPt + Len(replaceWith$) - 1
                                    End If
                                End If
                            Loop
                        End If
                    Next prp
                If abort Then Exit For
                Next ctl
            If abort Then Exit For
            wk$ = obj.name
            If Found Then
                hit = hit + 1
                DoCmd.Close acReport, wk$, acSaveYes
                Else
                DoCmd.Close acReport, wk$
                End If
            End If
        Next obj
    Set dbs = Nothing
    wk$ = "Examined: " & CStr(sought) & "  Found: " & CStr(hit)
    If abort Then wk$ = wk$ & "  (** aborted **)"
    If chn <> fig_NoLogFile Then
        Feedback True, chn, "", wk$
        Close #chn
        wk$ = wk$ & "  Written to: " & pth$
        End If
    Debug.Print wk$
    End Sub

Public Sub Feedback(ByRef Found As Boolean, ByVal chn%, ByVal objName$, ByVal txt$)
                    
'   Display or write to file the occurrence of a 'hit'

    If chn = fig_NoLogFile Then
        If Not Found Then Debug.Print "**** " & objName ' first hit for doc
        Debug.Print txt$
        Else
        If Not Found Then Print #chn, "**** " & objName ' first hit for doc
        Print #chn, txt$
        End If
    Found = True
    End Sub

Function ConfirmedChange(ByVal msg$ _
                       , ByVal prpVal$ _
                       , ByVal mtchPt% _
                       , ByVal searchFor$ _
              , Optional ByRef Cancel As Boolean = False _
              , Optional ByRef Confirm As Boolean _
              ) As Boolean
'   Try to put up msg that highlights where
'   in the string the change is to occur ...
'   ... to do this effectively, one has to mess with
'   fonts and proportional character sizes and stuff
'   ugh, nasty

    Dim wk$, msgTxt$, chSpc%, chMkr%, lnfSrch%, lnfMkr%, resp As VbMsgBoxResult
'   should check prpVal up to mtchPt to see if there is a line break
'   ... one day
    chSpc = TextWidthInSpaces(Left$(prpVal, mtchPt - 1))
    lnfSrch% = TextWidthInSpaces(searchFor)
    ' the 'clever' way to do this didn't work out, so the
    ' 2/3 factor below is an empirical fudge
    chMkr% = Int(Len(searchFor) * 2 / 3)
    
    msgTxt$ = msg$ & vbCrLf _
            & "     Value: " & prpVal$ & vbCrLf _
            & "     guide: " & Space$(chSpc) & String$(chMkr%, "^") & vbCrLf
    resp = MsgBox(msgTxt$, vbQuestion + vbYesNoCancel + vbDefaultButton1 _
                                                     , "Confirm change" _
                                                     )
    ' *^* when / if I decide to use a 4-button msgbox ....
    ' ... If resp = dont_ask_for_more_confirmation Then Confirm = False
    
    Cancel = (resp = vbCancel)
    ConfirmedChange = (resp = vbYes)
    End Function

Function TextWidthInSpaces(txt$) As Integer
'   orig code from MSDN, then hacked

'   return the number of spaces (proportional font) it would
'   take to occupy the same screen width as the given string

    Dim TxtWidth As Long, SpaceWidth As Long
    Dim hWnd As Long, hDC As Long
    Dim lpSIZE As Size
    Dim RetVal As Long

    hWnd = GetFocus()
    hDC = GetDC(hWnd)
    
    ' Compute the width of the text.
    RetVal = GetTextExtentPoint(hDC, txt, Len(txt), lpSIZE)
    winApi_ConvertPixelsToTwips lpSIZE.cx, lpSIZE.cy
    TxtWidth = lpSIZE.cx
    
    ' Compute width of a space.
    RetVal = GetTextExtentPoint(hDC, " ", 1, lpSIZE)
    winApi_ConvertPixelsToTwips lpSIZE.cx, lpSIZE.cy
    SpaceWidth = lpSIZE.cx
    
    TextWidthInSpaces = (TxtWidth / SpaceWidth)
    End Function

Sub winApi_ConvertPixelsToTwips(ByRef Xpix, ByRef Ypix)
'*******************************************
'Purpose:   Accepts two pixel measurements
'           returns two twips measurements
'Author:    CD
'Date:      April 1999
'Called by: Form fit checkers and routines playing with fonts/msg lengths
'*******************************************

    Dim Xppi As Long, Yppi As Long
    Dim lngDC As Long, RetVal As Long

    ' Retrieve the current number of pixels per inch, which is resolution-dependent.

    lngDC = GetDC(0)
    Xppi = GetDeviceCaps(lngDC, LOGPIXELSX)
    Yppi = GetDeviceCaps(lngDC, LOGPIXELSY)
    RetVal = ReleaseDC(0, lngDC)
    
    ' Compute and return the measurements in twips.
    Xpix = (Xpix * 1# / Xppi) * fig_TwipsPerInch
    Ypix = (Ypix * 1# / Yppi) * fig_TwipsPerInch
    End Sub

Public Function QuoteProtect(str) As String
'*******************************************
'Purpose:   put single quote marks (double-quotes) around a string,
'           ... and check for/duplicate quote mark in the string
'
'Author:    CD
'Date:      Jan 1996, 22:15
'Called by: General code,
'               typically priming Form/Control properties,
'               wrapping long path names,
'               code that generates code, etc etc
'*******************************************
    
    Dim qq$, qu$, i%, j%, wk$, txt$
    qu$ = """"                           ' aka Chr$(34)
    qq$ = qu$ & qu$: i = 0: wk = ""
    If Not IsNull(str) Then
        txt = str
        Do
            j = i + Len(qu$)
            i = InStr(j, txt, qu$)
            If i = 0 Then
                wk = wk & Mid$(txt, j)
                Exit Do
                End If
            wk = wk$ & Mid$(txt, j, i - j) & qq$    ' double up on embedded quote
            Loop
            End If
    QuoteProtect = qu$ & wk & qu$       ' wrap response in quotes
    End Function

Public Function SQL_QuoteProtect(str) As String
'*******************************************
'Purpose:   put single apostrophes (single-quote) around a string,
'           ... and check for/duplicate apostrophe in the string
'
'Author:    CD
'Date:      Jan 1996, 22:15
'Called by: General code, typically building SQL statement or criteria, DLookup etc
'*******************************************
    
    Dim qq$, qu$, i%, j%, wk$, txt$
    qu$ = "'"                           ' aka Chr$(39)
    qq$ = qu$ & qu$: i = 0: wk = ""
    If Not IsNull(str) Then
        txt = str
        Do
            j = i + Len(qu$)
            i = InStr(j, txt, qu$)
            If i = 0 Then
                wk = wk & Mid$(txt, j)
                Exit Do
                End If
            wk = wk$ & Mid$(txt, j, i - j) & qq$    ' double up on embedded quote
            Loop
            End If
    SQL_QuoteProtect = qu$ & wk & qu$       ' wrap response in quotes
    End Function

Public Function Made_New_Query(qryName$, sql$) As Boolean
'*******************************************
'Purpose:   Create new query in current db
'Author:    CD
'Date:      april 1999
'Called by: general system use
'*******************************************

    On Error Resume Next
    Made_New_Query = False
    Dim db As DAO.Database, qryDef As DAO.QueryDef
    Set db = CurrentDb()
    db.QueryDefs.Refresh
    db.QueryDefs.Delete qryName$                    '   remove any earlier version
    On Error GoTo mnq_err
    Set qryDef = db.CreateQueryDef(qryName$)
    qryDef.sql = sql$
    Made_New_Query = True

mnq_exit:
    On Error Resume Next
    qryDef.Close
    db.Close
    Set qryDef = Nothing
    Set db = Nothing
    Exit Function
mnq_err:
#If used_by_CD Then
    System_Error_Log "Made_New_Query", qryName$ & " %% " & sql$
#End If
    Resume mnq_exit
    End Function

Public Function TempQuery_NewName(core$) As String

    '   invent a query name that is not in use/does not exist
    
    Dim i%, wk$, qn$
    Dim db As DAO.Database
    Set db = CurrentDb()
    If core$ = "" Then
        wk = figTempFile_CoreName
        Else
        wk = core$
        End If
    On Error Resume Next                            '   .Delete method gives error if object not found
    db.QueryDefs.Refresh
    i = 0
    Do
        i = i + 1
        qn = wk & Right$(str$(100 + i), 2)
        db.QueryDefs.Delete qn                      '   delete object - if it exists
        If Err = 0 Then Exit Do                     '   object of the given name has just been deleted ... so fine
        If Err = err_ItemNotInThisCollection Then
                                                    '   name not found in this collection ... so fine
            Exit Do
            End If
        '   if object not deleted, then it must be 'in use' or some
        '   other problem, so build new name
        Err.Clear
        Loop
    TempQuery_NewName = qn
    End Function

Public Function TempQuery_DeletedAll(ByVal core$) As Boolean

    '   delete all queries with a temporary-style file name
    
    Dim i%, wk$
    Dim db As DAO.Database, qryD As DAO.QueryDef
    Set db = CurrentDb()
    If core$ = "" Then
        wk = figTempFile_CoreName
        Else
        wk = core$
        End If
    On Error Resume Next        '   .Delete method gives error if object not found
    db.QueryDefs.Refresh
    Dim nms$()
    i = 0
    For Each qryD In db.QueryDefs
        If 1 = InStr(qryD.name, wk$) Then
            i = i + 1: ReDim nms$(i)
            nms$(i) = qryD.name
            End If
        Next
    If i <> 0 Then
        i = 0
        Do
            i = i + 1: If i > UBound(nms$) Then Exit Do
            db.QueryDefs.Delete nms$(i)                 '   delete object - if it exists
            If Err <> 0 Then
                If Err <> err_ItemNotInThisCollection Then
                        '   name not found in this collection ... so fine, already gone
                '   if object not deleted, then it must
                '   be 'in use' or some other problem, so check for next
                    If vbYes <> MsgBox("Err: " & Err.Number & " - " & Err.Description _
                                       & vbCrLf & vbCrLf & "Query: " & nms$(i) _
                                       & vbCrLf & vbCrLf & "Continue?" _
                                     , vbYesNo + vbQuestion + vbDefaultButton2 _
                                     , "Delete Temp Query") Then
                        Exit Do
                        End If
                    End If
                End If
            Err.Clear
            Loop
        End If
    Set db = Nothing
    TempQuery_DeletedAll = True
    End Function
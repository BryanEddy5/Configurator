Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Export2XLS
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Export recordset to Excel
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sQuery    : Name of the table, or SQL Statement to be used to export the records
'             to Excel
'
' Usage:
' ~~~~~~
' Export2XLS "qryCustomers"
' Call Export2XLS("qryCustomers")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-Apr-18                 Initial Release
' 2         2015-May-01                 Header Clarifications
'---------------------------------------------------------------------------------------
Sub Export2XLS(ByVal sQuery As String, Optional ByVal sFileName As String, Optional ByVal sDirectory As String)
    Dim oExcel          As Object
    Dim oExcelWrkBk     As Object
    Dim oExcelWrSht     As Object
    Dim bExcelOpened    As Boolean
    Dim db              As DAO.Database
    Dim rs              As DAO.Recordset
    Dim iCols           As Integer
    Const xlCenter = -4108
    
    DoCmd.Hourglass True

    If IsNull(sDirectory) Or sDirectory = "" Then
        sDirectory = Environ("USERPROFILE") & "\Desktop\"
    End If
 
    'Start Excel
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")    'Bind to existing instance of Excel
 
    If Err.Number <> 0 Then    'Could not get instance of Excel, so create a new one
        Err.Clear
        On Error GoTo Error_Handler
        Set oExcel = CreateObject("Excel.Application")
        bExcelOpened = False
    Else    'Excel was already running
        bExcelOpened = True
    End If
    On Error GoTo Error_Handler
    oExcel.ScreenUpdating = False
    oExcel.Visible = False   'Keep Excel hidden until we are done with our manipulation
    Set oExcelWrkBk = oExcel.Workbooks.Add()    'Start a new workbook
    Set oExcelWrSht = oExcelWrkBk.Sheets(1)
 
    'Open our SQL Statement, Table, Query
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sQuery, dbOpenSnapshot)
    With rs
        If .RecordCount <> 0 Then
            'Build our Header
            For iCols = 0 To rs.Fields.Count - 1
                oExcelWrSht.Cells(1, iCols + 1).value = rs.Fields(iCols).name
            Next
            With oExcelWrSht.Range(oExcelWrSht.Cells(1, 1), _
                                   oExcelWrSht.Cells(1, rs.Fields.Count))
                .Font.Bold = True
                .Font.ColorIndex = 2
                .Interior.ColorIndex = 1
                .HorizontalAlignment = xlCenter
            End With
            oExcelWrSht.Range(oExcelWrSht.Cells(1, 1), _
                              oExcelWrSht.Cells(1, rs.Fields.Count)).Columns.AutoFit    'Resize our Columns based on the headings
            'Copy the data from our query into Excel
            oExcelWrSht.Range("A2").CopyFromRecordset rs
            oExcelWrSht.Range("A1").Select  'Return to the top of the page
        Else
            MsgBox "There are no records returned by the specified queries/SQL statement.", vbCritical + vbOKOnly, "No data to generate an Excel spreadsheet with"
            GoTo Error_Handler_Exit
        End If
    End With
    
    
        oExcelWrkBk.SaveAs sDirectory & sFileName
        'oExcelWrkBk.Close True, sfileName 'Save and close the generated workbook
 
    '    'Close excel if is wasn't originally running
    '    If bExcelOpened = False Then
    '        oExcel.Quit
    '    End If
    
        
 
Error_Handler_Exit:
    On Error Resume Next
    oExcel.Visible = True   'Make excel visible to the user
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set oExcelWrSht = Nothing
    Set oExcelWrkBk = Nothing
    oExcel.ScreenUpdating = True
    Set oExcel = Nothing
    DoCmd.Hourglass False
    Exit Sub
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Export2XLS" & vbCrLf & _
           "Error Description: " & Err.Description _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Sub
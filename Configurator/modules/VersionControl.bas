Option Explicit
Option Compare Database

Public Sub DocDatabase()
 '====================================================================
 ' Name:    DocDatabase
 ' Purpose: Documents the database to a series of text files
 '
 ' Author:  Arvin Meyer
 ' Date:    June 02, 1999
 ' Comment: Uses the undocumented [Application.SaveAsText] syntax
 '          To reload use the syntax [Application.LoadFromText]
 '      Modified to set a reference to DAO 8/22/2005
 '====================================================================
On Error GoTo Err_DocDatabase
Dim dbs As DAO.Database
Dim cnt As DAO.Container
Dim doc As DAO.Document
Dim i As Integer
Dim path$

path = Environ("USERPROFILE") & "\Document\"

Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

Set cnt = dbs.Containers("Forms")
For Each doc In cnt.Documents
    Application.SaveAsText acForm, doc.name, "D:\Document\" & doc.name & ".txt"
Next doc

Set cnt = dbs.Containers("Reports")
For Each doc In cnt.Documents
    Application.SaveAsText acReport, doc.name, "D:\Document\" & doc.name & ".txt"
Next doc

Set cnt = dbs.Containers("Scripts")
For Each doc In cnt.Documents
    Application.SaveAsText acMacro, doc.name, "D:\Document\" & doc.name & ".txt"
Next doc

Set cnt = dbs.Containers("Modules")
For Each doc In cnt.Documents
    Application.SaveAsText acModule, doc.name, "D:\Document\" & doc.name & ".txt"
Next doc

For i = 0 To dbs.QueryDefs.Count - 1
    Application.SaveAsText acQuery, dbs.QueryDefs(i).name, "D:\Document\" & dbs.QueryDefs(i).name & ".txt"
Next i

Set doc = Nothing
Set cnt = Nothing
Set dbs = Nothing

Exit_DocDatabase:
    Exit Sub


Err_DocDatabase:
    Select Case Err

    Case Else
        MsgBox Err.Description
        Resume Exit_DocDatabase
    End Select

End Sub
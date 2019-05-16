Option Explicit
Option Compare Database


Function TurnOffSubDataSheets()
Dim MyDB As DAO.Database
Dim MyProperty As DAO.Property

Dim propName As String
Dim propType As Integer
Dim propVal As String
Dim i
Dim intChangedTables%


Set MyDB = CurrentDb

propName = "SubDataSheetName"
propType = 10
propVal = "[NONE]"

On Error Resume Next

For i = 0 To MyDB.TableDefs.Count - 1

If (MyDB.TableDefs(i).Attributes And dbSystemObject) = 0 Then

If MyDB.TableDefs(i).Properties(propName).value < propVal Then
MyDB.TableDefs(i).Properties(propName).value = propVal
intChangedTables = intChangedTables + 1
End If

If Err.Number = 3270 Then
Set MyProperty = MyDB.TableDefs(i).CreateProperty(propName)
MyProperty.Type = propType
MyProperty.value = propVal
MyDB.TableDefs(i).Properties.Append MyProperty
Else
If Err.Number < 0 Then
MsgBox "Error: " & Err.Number & " on Table " _
& MyDB.TableDefs(i).name & "."
MyDB.Close
Exit Function
End If
End If

End If
Next i

MsgBox "The " & propName & _
" value for all non-system tables has been updated to " & propVal & "."

End Function



Function TableTurnOffDataSubSheet(strTable As String)
Dim MyDB As DAO.Database
Dim MyProperty As DAO.Property

Dim propName As String
Dim propType As Integer
Dim propVal As String
Dim proptable$, intChangedTables%

Set MyDB = CurrentDb

propName = "SubDataSheetName"
propType = 10
propVal = "[NONE]"
proptable = strTable

On Error Resume Next


If (MyDB.TableDefs(strTable).Attributes And dbSystemObject) = 0 Then

If MyDB.TableDefs(strTable).Properties(propName).value < propVal Then
MyDB.TableDefs(strTable).Properties(propName).value = propVal
intChangedTables = intChangedTables + 1
End If

If Err.Number = 3270 Then
Set MyProperty = MyDB.TableDefs(strTable).CreateProperty(propName)
MyProperty.Type = propType
MyProperty.value = propVal
MyDB.TableDefs(strTable).Properties.Append MyProperty
Else
If Err.Number < 0 Then
MsgBox "Error: " & Err.Number & " on Table " _
& MyDB.TableDefs(strTable).name & "."
MyDB.Close
Exit Function
End If
End If

End If


'MsgBox "The " & propName & _
'" value for all non-system tables has been updated to " & propVal & "."

End Function


Function CreateSimpleShortcutMenu(frm As String)


    Dim cmbRightClick As Office.CommandBar
 
 ' Create the shortcut menu.
    Application.CommandBars(frm).Delete
    Set cmbRightClick = CommandBars.Add(frm, msoBarPopup, False, True)
     
    With cmbRightClick
        ' Add the Find command.
        .Controls.Add msoControlButton, 141, , , True
         
        ' Start a new grouping and add the Sort Ascending command.
        .Controls.Add(msoControlButton, 210, , , True).BeginGroup = True
         
        ' Add the Sort Descending command.
        .Controls.Add msoControlButton, 211, , , True
         
        ' Start a new grouping and add the Remove Filer/Sort command.
        .Controls.Add(msoControlButton, 605, , , True).BeginGroup = True
         
        ' Add the Filter by Selection command.
        .Controls.Add msoControlButton, 640, , , True
         
        ' Add the Filter Excluding Selection command.
        .Controls.Add msoControlButton, 3017, , , True
         
        ' Add the Between... command.
        .Controls.Add msoControlButton, 10062, , , True
    End With
 
Set cmbRightClick = Nothing


End Function


Function DoFolder(Folder)


    Dim SubFolder
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
    Next
    Dim File
    For Each File In Folder.Files
        ' Operate on each file
    Next
End Function

    Sub CallingProcedure()
        Dim A As Long
        Dim b As Long
        A = 123
        b = 456
        Debug.Print "BEFORE CALL = X: " & CStr(A), "Y: " & CStr(b)
        CalledProcedure x:=A, Y:=b
        Debug.Print "AFTER CALL =  X: " & CStr(A), "Y: " & CStr(b)
    End Sub
    
    Sub CalledProcedure(ByRef x As Long, ByVal Y As Long)
        x = 321
        Y = 654
    End Sub
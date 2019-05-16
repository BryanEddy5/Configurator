Option Compare Database
Option Explicit

Function TurnOffSubDataSheets()
Dim MyDB As DAO.Database
Dim MyProperty As DAO.Property

Dim propName As String
Dim propType As Integer
Dim propVal As String
Dim i As Variant
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



Function TurnOffTableDataSheet()
Dim MyDB As DAO.Database
Dim MyProperty As DAO.Property

Dim propName As String
Dim propType As Integer
Dim propVal As String
Dim proptable As String
Dim intChangedTables%

Set MyDB = CurrentDb

propName = "dbo_RibbonSafetyStock"
propType = 10
propVal = "[NONE]"
proptable = "dbo_RibbonFgItems"

On Error Resume Next



If (MyDB.TableDefs(proptable).Attributes And dbSystemObject) = 0 Then

If MyDB.TableDefs(proptable).Properties(propName).value < propVal Then
MyDB.TableDefs(proptable).Properties(propName).value = propVal
intChangedTables = intChangedTables + 1
End If

If Err.Number = 3270 Then
Set MyProperty = MyDB.TableDefs(proptable).CreateProperty(propName)
MyProperty.Type = propType
MyProperty.value = propVal
MyDB.TableDefs(proptable).Properties.Append MyProperty
Else
If Err.Number < 0 Then
MsgBox "Error: " & Err.Number & " on Table " _
& MyDB.TableDefs(proptable).name & "."
MyDB.Close
Exit Function
End If
End If

End If



End Function
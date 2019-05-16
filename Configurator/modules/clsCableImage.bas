Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Function Load(ImageID As String) As Boolean


On Error GoTo HandleError:

Dim sQry As String
Dim rst As New ADODB.Recordset

sQry = "SELECT * FROM CutSheet.Images WHERE ImageID = " & ImageID


Set rst = Core_DataAccess.GetAdodbRecordset(sQry)
 
    Load = False


    With rst
        If .EOF Then
            GoTo Done
        End If
        
        
            ' Note use of Nz function in case of Null values in the table

        .Close
    End With
    Load = True
    

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function
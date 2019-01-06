VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCableImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Function Load(ImageID As String) As Boolean


On Error GoTo HandleError:

Dim cnn As New ADODB.Connection
Set cnn = New ADODB.Connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = cPremiseServerConnection
cnn.Open provstr
sQry = "SELECT * FROM CutSheet.Images WHERE ImageID = " & ImageID


rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
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

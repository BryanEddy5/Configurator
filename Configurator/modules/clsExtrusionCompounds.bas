Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Color As String
Public JacketMaterial As String
Public ColorChip As String
Public Color_Chip_Concentration As Double
Public Additive As String
Public AdditivePercentage As Double
Public Additive2 As String
Public AdditivePercentage2 As Double
Public JacketPercentage As Double
Public SpecificGravity As Double


Public Function Load(item As String) As Boolean


On Error GoTo HandleError:

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset



' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
sQry = "SELECT * FROM vExtrusionMaterialsItem WHERE Oracle = '" & item & "'"


rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 



    Load = False



    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If

        ' Note use of Nz function in case of Null values in the table
        Me.Color = Nz(!Color)
        Me.JacketMaterial = Nz(!JacketMaterial)
        Me.ColorChip = Nz(!ColorChip)
        Me.Color_Chip_Concentration = Nz(!Color_Chip_Concentration)
        Me.Additive = Nz(!Additive)
        Me.AdditivePercentage = Nz(!AdditivePercentage)
        Me.Additive2 = Nz(!Additive2)
        Me.AdditivePercentage2 = Nz(!AdditivePercentage2)
        Me.JacketPercentage = Nz(!JacketPercentage)
        Me.SpecificGravity = Nz(!SpecificGravity)
        
        If IsNull(Me.JacketMaterial) Then
            MsgBox "Jacket material density is missing", vbCritical
        End If

        .Close
    End With
    Load = True

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading compound information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function

Public Function CompoundCheck(Compound As String) As Boolean


On Error GoTo HandleError:

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset



' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
sQry = "SELECT * FROM vCompoundsInOracle WHERE Compound = '" & Compound & "'"


rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 



    CompoundCheck = False



    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If
        .Close
    End With
    CompoundCheck = True

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while checking compound information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function
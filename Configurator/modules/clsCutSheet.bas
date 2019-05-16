Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Public ItemNo As String
Public NominalOD As Double
Public TensileRatingLongTerm_N As Integer
Public NumFillers As Integer
Public NumSubs As Integer
Public Description As String
Public Operation_C As String
Public StorageShipping_C As String
Public Installation_C As String
Public FiberCount As Integer
Public Location As Integer
Public CableFamily As String
Public Base As String
Public Commercial_Approval As Boolean
Public Technical_Approval As Boolean
Public TensileRatingShortTerm_N As Integer
Public BaseID As Integer
Public Requested As Boolean
Private mlItem As String
'Public ImageGroupID%





Public Function Load(item As String) As Boolean

'''Instantiates the Class'''


On Error GoTo HandleError:

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset

mlItem = item


' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
sQry = "SELECT * FROM vCutSheet WHERE Oracle = '" & item & "'"


rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 


    Load = False




    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If

        ' Note use of Nz function in case of Null values in the table
        Me.ItemNo = Nz(!DesignCode)
        Me.NominalOD = Nz(!NominalOD)
        Me.TensileRatingLongTerm_N = Nz(!TensileRatingLongTerm_N)
        Me.NumFillers = Nz(!NumSubFillers)
        Me.NumSubs = Nz(!NumSubPositions)
        Me.Description = Nz(!SalesForceDescription)
        Me.Operation_C = Nz(!Operation_C)
        Me.StorageShipping_C = Nz(!StorageShipping_C)
        Me.Installation_C = Nz(!Installation_C)
        Me.FiberCount = Nz(!FiberCount)
        Me.Location = Nz(!Location)
        Me.CableFamily = Nz(!CableFamily)
        Me.Base = Nz(!Base)
        Me.Technical_Approval = Nz(!Technical_Approval)
        Me.Commercial_Approval = Nz(!Commercial_Approval)
        Me.TensileRatingShortTerm_N = Nz(!TensileRatingShortTerm_N)
        Me.BaseID = Nz(!BaseID)
        Me.Requested = Nz(!Requested)
        'Me.ImageGroupID = Nz(!ImageGroupID)
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


Property Get CutSheetApproval() As Boolean      ' Read only

        
        CutSheetApproval = Commercial_Approval * Technical_Approval

    
End Property

Property Get getCutSheetApproval() As Boolean      ' Read only

        
        getCutSheetApproval = Commercial_Approval * Technical_Approval

    
End Property







'Property Get NominalOD() As Long      ' Read only
'    NominalOD = mlNominalOD
'End Property
'
'Property Get TensionLongTerm() As Long      ' Read only
'    TensionLongTerm = mlTensionLongTerm
'End Property
'
'Property Get NumFillers() As Long      ' Read only
'    NumFillers = mlNumFillers
'End Property
'
''Property Get Description() As Long      ' Read only
''    Description = Me.Description
''End Property
'
'Property Get Operation_C() As Long      ' Read only
'    Operation_C = mlOperation_C
'End Property
'
'Property Get StorageShipping() As Long      ' Read only
'    StorageShipping = mlStorageShipping_C
'End Property
'
'Property Get Installation_C() As Long      ' Read only
'    Installation_C = mlInstallation_C
'End Property
'
'
'Property Get FiberCount() As Long      ' Read only
'    FiberCount = mlFiberCount
'End Property
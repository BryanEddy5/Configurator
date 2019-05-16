Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Fiber As String, Micron As String, BIF As String, OM As String, SpecialNotes As String, FiberType As String, Mode As String, Vendor As String, ITU As String, mlFiberLetter As String
Private mlFiberCount As Integer
Private mlVendorSpecific As Boolean



Public Function Load(FiberType As String) As Boolean


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
sQry = "SELECT * FROM dbo.vFiber WHERE Fiber = '" & FiberType & "'"

Me.Fiber = FiberType

rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
    Load = False


    With rst
        If .EOF Then
            GoTo Done
        End If
            Me.Micron = Nz(![SM/MM])
            Me.BIF = Nz(!BIF)
            Me.OM = Nz(![OM#])
            Me.SpecialNotes = Nz(!SpecialNotes)
            Me.FiberType = Nz(!Fiber_Type)
            Me.Mode = Nz(!Mode)
            Me.Vendor = Nz(!Vendor)
            Me.ITU = Nz(!ITU)
        
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

Public Property Let setFiberCount(FiberCount As Integer)
mlFiberCount = FiberCount

End Property

Public Property Get getFiberCount() As Integer

getFiberCount = mlFiberCount

End Property

Public Property Let setFiberLetter(FiberLetter As String)
mlFiberLetter = FiberLetter

End Property

Public Property Get getFiberLetter() As String

    getFiberLetter = mlFiberLetter

End Property
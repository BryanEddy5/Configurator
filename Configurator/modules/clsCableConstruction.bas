Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Public Base, ConstructionDescription As String, JacketMaterial As String, DesignCode As String, TBType As String, subBase As String, fillerBase As String, fillerOD As String, StandardOperation  As String
Public NumPositions As Integer, NumFillers As Integer, Location As Integer, TBTypeID As Integer, FibersPerBundle As Integer, BaseID As Integer, ProductID As Integer, SubFiberCount As Integer, subFiberPerBundle As Integer, DesignTypeID As Integer, FiberCount As Integer, SetupID As Integer
Public Active As Boolean, ReleasedDesign As Boolean, TightBuffered As Boolean, RibbonHighCure As Boolean, IsCommoned As Boolean
Public SubOD#, NominalOD@


Public Function Load(Base As String) As Boolean


On Error GoTo HandleError:


Dim sQry As String
Dim rst As New ADODB.Recordset


sQry = "SELECT * FROM dbo.[vCableConstruction] WHERE Base = '" & Base & "'"


Set rst = Core_DataAccess.GetAdodbRecordset(sQry)
 
    Load = False


    With rst
        If .EOF Then
            GoTo Done
        End If
        
        
        
        
    
            ' Note use of Nz function in case of Null values in the table
            Me.Base = Nz(!Base)
            Me.ConstructionDescription = Nz(!ConstructionDescription)
            Me.JacketMaterial = Nz(!JacketMaterial)
            Me.TBType = Nz(!TBType)
            Me.NumPositions = Nz(!NumPositions)
            Me.NumFillers = Nz(!NumFillers)
            Me.Location = Nz(!Location)
            Me.TBTypeID = Nz(!TBTypeID)
            Me.FibersPerBundle = Nz(!FibersPerBundle)
            Me.Active = Nz(!Active)
            Me.ReleasedDesign = Nz(!ReleasedDesign)
            Me.TightBuffered = Nz(!TightBuffered)
            Me.BaseID = Nz(!BaseID)
            Me.ProductID = !ProductID
            Me.DesignTypeID = Nz(!DesignTypeID)
            Me.FiberCount = Nz(!FiberCount)
            Me.DesignCode = Nz(!DesignCode)
            Me.SetupID = Nz(!SetupID)
            Me.StandardOperation = Nz(!StandardOperation)
            Me.NominalOD = Nz(!OD)
            Me.IsCommoned = Nz(!IsCommoned)
        .Close
    End With
    Load = True
    
Call getSubUnit
Call getFiller
Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function


Private Function getSubUnit() As Boolean


On Error GoTo HandleError:

If IsNull(ProductID) Then
    Exit Function
End If


Dim sQry As String
Dim rst As New ADODB.Recordset

sQry = "SELECT * FROM vSubunitConstruction WHERE ProductID = " & Nz(SetupID) & " AND subFiller = 0"

Set rst = Core_DataAccess.GetAdodbRecordset(sQry)
 
    getSubUnit = False


    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If
            
            ' Note use of Nz function in case of Null values in the table
            Me.subBase = Nz(!Base)
            Me.SubOD = Nz(!NominalOD)
            'Me.subFiberCount = Nz(!FiberCount)
            'Me.FibersPerBundle = Nz(!FibersPerBundle)
        .Close
    End With
    getSubUnit = True

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function

Private Function getFiller() As Boolean


On Error GoTo HandleError:

If IsNull(ProductID) Then
    Exit Function
End If


Dim sQry As String
Dim rst As New ADODB.Recordset

sQry = "SELECT * FROM vSubunitConstruction WHERE ProductID = " & Nz(SetupID) & " AND subFiller <> 0"

Set rst = Core_DataAccess.GetAdodbRecordset(sQry)
 
    getFiller = False


    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If
            
            ' Note use of Nz function in case of Null values in the table
            Me.fillerBase = Nz(!Base)
            Me.fillerOD = Nz(!NominalOD)
            'Me.subFiberCount = Nz(!FiberCount)
            'Me.FibersPerBundle = Nz(!FibersPerBundle)
        .Close
    End With
    getFiller = True

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function


Public Function SetupExist(StandardOperation As String) As Boolean


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
sQry = "SELECT * FROM dbo.tblCableConstructions WHERE StandardOperation = '" & StandardOperation & "'"


rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
    SetupExist = False


    With rst
        If .EOF Then
            GoTo Done
        End If
    
            ' Note use of Nz function in case of Null values in the table
            Me.BaseID = Nz(!BaseID)
            'Me.StandardOperation = Nz(!StandardOperation)
        .Close
    End With
    SetupExist = True
    

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function


Public Function BaseExist(Base As String) As Boolean


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
sQry = "SELECT * FROM dbo.tblCableConstructionReferences WHERE Base = '" & Base & "'"


rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
    BaseExist = False


    With rst
        If .EOF Then
            GoTo Done
        End If
    
            ' Note use of Nz function in case of Null values in the table
            Me.SetupID = Nz(!SetupID)
            'Me.StandardOperation = Nz(!StandardOperation)
        .Close
    End With
    BaseExist = True
    

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function
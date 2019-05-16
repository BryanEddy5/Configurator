Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private mlUserName As String
Private mlResponsibilities As New Scripting.Dictionary
Private mlUserID As Integer, mlUserEmail As String

Public Function Load() As Boolean

Dim rst As New ADODB.Recordset
Dim qry As String

mlUserName = Core_VersionCheck.GetCurrentUser

Load = False

qry = "SELECT UserID, UserEmail FROM tblConfiguratorUser WHERE [User] = '" & mlUserName & "'"

Set rst = Core_DataAccess.GetAdodbRecordset(qry)
 
    Load = False


    With rst
        If .EOF Then
            GoTo Done
        End If
    
            ' Note use of Nz function in case of Null values in the table
            mlUserID = Nz(!UserId)
            mlUserEmail = Nz(!userEmail)
            
            Call loadUserResponsibilities

        .Close
    End With
    Load = True
    


Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading user information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
End Function

Public Function getAdminEmailsSemiColonDelimited() As String


Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim rst As New ADODB.Recordset
Dim provstr As String
Dim qry As String
Dim i As Integer




' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr


qry = " SELECT * FROM Users.vUserResponsibility where [ResponsibilityID] = " & 12

rst.Open (qry), cnn, adOpenKeyset, adLockOptimistic


    With rst
        If .EOF Then
            MsgBox "Admin emails could not be found", vbCritical
            GoTo Done
        End If
        
        Do Until .EOF
            'mlResponsibilities.Add
            getAdminEmailsSemiColonDelimited = getAdminEmailsSemiColonDelimited & ";" & rst!userEmail
            rst.MoveNext
        Loop

        .Close
    End With


Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading admin email(s):" & vbCrLf & Err.Description, vbCritical
    Resume Done



End Function

Private Function loadUserResponsibilities() As Boolean

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim rst As New ADODB.Recordset
Dim provstr As String
Dim qry As String
Dim i As Integer

    mlResponsibilities.RemoveAll



' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr



qry = " SELECT * FROM Users.vUserResponsibility where [User] = '" & mlUserName & "'"

rst.Open (qry), cnn, adOpenKeyset, adLockOptimistic


    loadUserResponsibilities = False


    With rst
        If .EOF Then
            GoTo Done
        End If
        
        Do Until .EOF
            'mlResponsibilities.Add
            mlResponsibilities.Add !ResponsibilityID.value, Nz(!startUpForm.value, "0")
            .MoveNext
        Loop

        .Close
    End With
    loadUserResponsibilities = True

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading user responsibilities:" & vbCrLf & Err.Description, vbCritical
    Resume Done

End Function

Public Property Get getUserName() As String

getUserName = mlUserName

End Property


Public Property Get getUserEmail() As String

getUserEmail = mlUserEmail

End Property

Public Property Get getResponsibilities() As Scripting.Dictionary

Set getResponsibilities = mlResponsibilities

End Property

Public Property Get UserId() As String

UserId = mlUserID

End Property
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private mlApplicationDirectory As String, mlApplicationName As String, mlClientDirectory As String, mlUserDirectory As String
Public Enum ApplicationType
    Configurator = 1
    Scheduling
End Enum


Public Function Load(Application As ApplicationType) As Boolean


On Error GoTo HandleError:


Dim sQry As String
Dim rst As New ADODB.Recordset


sQry = "SELECT ApplicationDirectory, ApplicationName, ClientDirectory FROM ApplicationDirectory WHERE ApplicationID = " & Application & ""

Set rst = Core_DataAccess.GetAdodbRecordset(sQry)
 
    Load = False


    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If
    
            ' Note use of Nz function in case of Null values in the table
            mlApplicationDirectory = Nz(!ApplicationDirectory)
            mlApplicationName = Nz(!ApplicationName)
            mlClientDirectory = Nz(!ClientDirectory)

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


Property Get getApplicationDirectory() As String      ' Read only
        getApplicationDirectory = mlApplicationDirectory

End Property

Property Get getApplicationName() As String
    getApplicationName = mlApplicationName
    
End Property

Property Get getClientDirectory() As String
    getClientDirectory = Environ("USERPROFILE") & mlClientDirectory

End Property

Property Get getUserDirectory() As String
    getUserDirectory = Environ("USERPROFILE") & mlClientDirectory

End Property


Property Get getClientApplicationDirectory() As String
    getClientApplicationDirectory = Environ("USERPROFILE") & mlClientDirectory & mlApplicationName

End Property

Property Get getServerApplicationDirectory() As String
    getServerApplicationDirectory = mlApplicationDirectory & "\" & mlApplicationName

End Property

Property Get getClientApplicationBatFileDirectory()
    getClientApplicationBatFileDirectory = Environ("USERPROFILE") & "\Desktop\Configurator.bat"
End Property

Property Get getServerApplicationBatFileDirectory()
   getServerApplicationBatFileDirectory = mlApplicationDirectory & "\Configurator.bat"
End Property
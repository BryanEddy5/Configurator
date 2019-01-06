Attribute VB_Name = "VersionCheck"
Option Compare Database
Option Explicit
Const FOREIGN_KEY_CONSTRAINT As String = "The .+ statement conflicted with the FOREIGN KEY constraint"
Const FOREIGN_KEY_CONSTRAINT_MSG As String = "Database error: The value in [LABEL] does not exist in the related table"
Global sUser As Integer
Global iCutSheetApprover As Integer
Global sUserName As String
Dim ApplicationDirectory As New clsApplicationDirectory
Private Version As Double

Public Property Get getVersion() As Double

    getVersion = 6.58

End Property

Function VerifyVersion()
Dim start As Double

Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String
'Dim DateCreated As Date
Dim provstr As String

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = cPremiseServerConnection
cnn.Open provstr



On Error GoTo ErrorHandler

ApplicationDirectory.Load (Configurator) 'Initialize the application object with the data for Premise Configurator

Call UpdateUserLastLogin
Call RemoveOldConfig
Call CreateConfigFolder
Call MakeShortCut
Call getBatFile
Call UpdateCdiTables


If UserCheck Then 'if the user exists then run the login procedure else the frmUserEmail form will launch the login procedure

    Login.OpenRoleBasedForm

End If

start = Timer
While Timer < start + 2
    DoEvents
Wend

strSQL = "SELECT * FROM [tblConfiguratorVersion] ORDER BY ConfiguratorVersion DESC"


rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic



If VersionCheck.getVersion <> rst!ConfiguratorVersion Then

    MsgBox ("Please wait while the program updates"), vbCritical
    
    'Update CDI tables
    

    DoCmd.Hourglass True

    'copy Access file
    CreateObject("Scripting.FileSystemObject").CopyFile _
        ApplicationDirectory.getServerApplicationDirectory, ApplicationDirectory.getClientApplicationDirectory
    'allow enough time for file to completely copy before opening
    start = 0
    

    start = Timer
    While Timer < start + 3
        DoEvents
    Wend
    'load new version - SysCmd function gets the Access executable file path
    'Shell function requires literal quote marks in the target filename string argument, hence the quadrupled quote marks
    Shell SysCmd(acSysCmdAccessDir) & "MSAccess.exe " & """" & ApplicationDirectory.getClientApplicationDirectory & """", vbNormalFocus
    'close current file
    DoCmd.Close acForm, "frmSwitchboard"
    DoCmd.Hourglass False
    DoCmd.Quit

MsgBox "Load Complete"

End If

'Run commands after the update

DoCmd.ShowToolbar "Ribbon", acToolbarNo

If Application.Version >= 15 And Application.CommandBars("Ribbon").Height > 61 Then
    ''CommandBars.ExecuteMso "MinimizeRibbon"
    CommandBars.ExecuteMso "HideRibbon"
End If




rst.Close


ErrorExit:
Exit Function

ErrorHandler:
Resume Next



End Function

Function UserCheck() As Boolean

Dim cnn As ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String

On Error GoTo ErrorHandler

Set cnn = CurrentProject.Connection

sUserName = GetCurrentUser()

UserCheck = False


strSQL = "SELECT * FROM tblConfiguratorUser WHERE User = '" & sUserName & "' AND UserEmail IS NOT NULL"

rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic


If rst.EOF = True Then

    DoCmd.OpenForm "frmUserEmail"
    
    MsgBox "Please add your AFL Global email address for notifications", vbCritical
    
    Call AccessVersionCheck
    
    Exit Function

End If

UserCheck = True
Call AccessVersionCheck

sUser = rst!UserTypeID
iCutSheetApprover = rst!CutSheetApprover
'UserCheck = rst!UserTypeID


rst.Close
cnn.Close

ErrorHandler:
Exit Function


End Function



Public Function ItemDateCreated(Oracle As String) As Boolean

' Initialize variables.
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String
'Dim DateCreated As Date
Dim provstr As String

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = cPremiseServerConnection
cnn.Open provstr


strSQL = "SELECT cast([DateCreated] as date) as CreationDate, cast(getdate() as date) as CurrentDate "
strSQL = strSQL & "FROM [Basic Product Construction] "
strSQL = strSQL & "WHERE [New Oracle Part #] = '" & Oracle & "'"

rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
'
'DateCreated = rst!CreationDate
'DateCreated = DateValue(DateCreated)
'If rst!CreationDate = Date Then
If rst!CreationDate = rst!CurrentDate Then
    ItemDateCreated = True
Else
ItemDateCreated = False
End If

rst.Close
cnn.Close

End Function

Function IsMember(strDomain As String, strGroup _
  As String, strMember As String) As Boolean
  Dim grp As Object
  Dim strPath As String

  strPath = "WinNT://" & strDomain & "/"
  Set grp = GetObject(strPath & strGroup & ",group")
  IsMember = grp.IsMember(strPath & strMember)
End Function



Public Property Get GetCurrentUser() As String
    GetCurrentUser = Environ("USERNAME")
End Property

Function GetCurrentDomain() As String
    GetCurrentDomain = Environ("USERDOMAIN")
End Function
Function getUserProfile() As String
    getUserProfile = Environ("USERPROFILE")
    
End Function
'
'Public Sub DocDatabase()
' '====================================================================
' ' Name:    DocDatabase
' ' Purpose: Documents the database to a series of text files
' '
' ' Author:  Arvin Meyer
' ' Date:    June 02, 1999
' ' Comment: Uses the undocumented [Application.SaveAsText] syntax
' '          To reload use the syntax [Application.LoadFromText]
' '====================================================================
'On Error GoTo Err_DocDatabase
'Dim dbs As Database
'Dim cnt As Container
'Dim doc As Document
'Dim i As Integer
'Dim Location As String
'
'Location = "C:\Users\eddybc\Desktop\Databases\Configurator Documents\"
'
'Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections
'
'Set cnt = dbs.Containers("Forms")
'For Each doc In cnt.Documents
'    Application.SaveAsText acForm, doc.name, Location & doc.name & ".txt"
'Next doc
'
'Set cnt = dbs.Containers("Reports")
'For Each doc In cnt.Documents
'    Application.SaveAsText acReport, doc.name, Location & doc.name & ".txt"
'Next doc
'
'Set cnt = dbs.Containers("Scripts")
'For Each doc In cnt.Documents
'    Application.SaveAsText acMacro, doc.name, Location & doc.name & ".txt"
'Next doc
'
'Set cnt = dbs.Containers("Modules")
'For Each doc In cnt.Documents
'    Application.SaveAsText acModule, doc.name, Location & doc.name & ".txt"
'Next doc
'
'For i = 0 To dbs.QueryDefs.Count - 1
'    Application.SaveAsText acQuery, dbs.QueryDefs(i).name, Location & dbs.QueryDefs(i).name & ".txt"
'Next i
'
'Set doc = Nothing
'Set cnt = Nothing
'Set dbs = Nothing
'
'Exit_DocDatabase:
'    Exit Sub
'
'
'Err_DocDatabase:
'Resume Next
''    Select Case Err
''
''    Case Else
''        MsgBox Err.Description
''        Resume Exit_DocDatabase
''    End Select
'
'End Sub

Public Function BaseID_DateCreated(BaseID As Integer) As Boolean

' Initialize variables.
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim strSQL As String
'Dim DateCreated As Date
Dim provstr As String

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = cPremiseServerConnection
cnn.Open provstr


strSQL = "SELECT cast([DateCreated] as date) as CreationDate, cast(getdate() as date) as CurrentDate "
strSQL = strSQL & "FROM [tblCableConstructions] "
strSQL = strSQL & "WHERE [BaseID] = '" & BaseID & "'"

rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic

'DateCreated = rst!DateCreated
'DateCreated = DateValue(DateCreated)
'If rst!CreationDate = Date Then
If rst!CreationDate = rst!CurrentDate Then
    BaseID_DateCreated = True
Else
BaseID_DateCreated = False
End If

rst.Close
cnn.Close

End Function



Public Sub MakeShortCut()
'''Create shortcut of the Configurator program on the user's desktop'''
Dim pathShortcut

pathShortcut = getUserProfile & "\Desktop\" & ApplicationDirectory.getApplicationName & ".lnk"
'If FileExists(pathShortcut) Then
'    Exit Sub
'End If


With CreateObject("WScript.Shell")
With .CreateShortcut(pathShortcut)
.TargetPath = ApplicationDirectory.getClientApplicationDirectory
.WindowStyle = 1
.Hotkey = "CTRL+SHIFT+G"
.IconLocation = ApplicationDirectory.getUserDirectory & "PremiseConfiguratorIcon.ico, 0"
.Description = "Premise Configurator"
.WorkingDirectory = ApplicationDirectory.getClientApplicationDirectory
.Save
End With
End With

End Sub

Private Sub CreateConfigFolder() 'Create Premise Configurator file and add the icon
'''Create folder to house the Configurator program'''
Dim Path As String
Dim localIconPath As String
localIconPath = ApplicationDirectory.getUserDirectory & "PremiseConfiguratorIcon.ico"

Path = ApplicationDirectory.getUserDirectory

If FileExists(Path) = False Then

    MkDir ApplicationDirectory.getUserDirectory

End If

CreateIcon (localIconPath)

End Sub

Private Sub RemoveOldConfig() ' iterate through all possible extensions of the Premise Configurator on the users desktop
'''Remove old versions of the Configurator program'''
Dim oldFile() As String
Dim File As String
Dim oldConfig As String
Dim fileExt As New Collection
Dim ext As Variant
oldFile = Split(ApplicationDirectory.getApplicationName, ".")

fileExt.Add ".mdb"
fileExt.Add ".accdb"
fileExt.Add ".accde"
fileExt.Add "." & oldFile(1)


File = oldFile(0)

For Each ext In fileExt


    oldConfig = getUserProfile & "\Desktop\" & File & ext
    
    
    If CurrentProject.FullName <> oldConfig Then
        DeleteFile (oldConfig)
    End If

Next

subRefreshDesktop


End Sub

Private Sub CreateIcon(localIconPath As String)

If FileExists(localIconPath) Then
    Exit Sub
End If

FileCopy ApplicationDirectory.getApplicationDirectory & "\PremiseConfiguratorIcon.ico", localIconPath

End Sub

Private Sub getBatFile()
'''get the Configurator program .bat file and add it to user's desktop'''

If FileExists(ApplicationDirectory.getClientApplicationBatFileDirectory) Then
    Exit Sub
End If

FileCopy ApplicationDirectory.getServerApplicationBatFileDirectory, ApplicationDirectory.getClientApplicationBatFileDirectory


End Sub

Private Sub UpdateCdiTables()
'''get the CDI Access tables and load them to the appropriate folder

Dim ClientPath As String
Dim RemotePath As String
RemotePath = ApplicationDirectory.getApplicationDirectory & "\CDITables.mdb"

ClientPath = "C:\Applications\CDITables.mdb"

If FileExists(ClientPath) = False Then

    MkDir "C:\Applications"

End If


FileCopy RemotePath, ClientPath


End Sub

Sub UpdateUserLastLogin()
'''Update the last login date by the user'''


Dim cnn As New ADODB.Connection
Set cnn = New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim provstr As String
Dim strStoredProcedure As String

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = cOracleExtractServerConnection
cnn.Open provstr


strStoredProcedure = " DECLARE @return_value INT EXEC  @return_value = [Users].[usp_UpdateUserLastLogin] @UserName = N'" & GetCurrentUser & "' SELECT  'ReturnValue' = @return_value "

rst.Open (strStoredProcedure), cnn, adOpenKeyset, adLockOptimistic


If rst!ReturnValue = 1 Then
    'do something if this procedure fails
End If

'Debug.Print "Success"


rst.Close
cnn.Close



 
End Sub




Private Sub AccessVersionCheck()
'''Check the version of the access the user is using'''
If Application.Version < 15 Then
    MsgBox "Please update to MS Access 2013 or newer to ensure all features work properly.", vbCritical
End If


End Sub

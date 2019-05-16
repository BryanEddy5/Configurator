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

    getVersion = 6.59

End Property

Function VerifyVersion()
    Dim start As Double

    Dim rst As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo ErrorHandler


    ApplicationDirectory.Load (Configurator)     'Initialize the application object with the data for Premise Configurator

    Call UpdateUserLastLogin
    Call AccessVersionCheck


    If UserCheck Then                            'if the user exists then run the login procedure else the frmUserEmail form will launch the login procedure

        Core_Login.OpenRoleBasedForm

    End If

    start = Timer
    While Timer < start + 2
        DoEvents
    Wend

    strSQL = "SELECT * FROM [tblConfiguratorVersion] ORDER BY ConfiguratorVersion DESC"


    Set rst = Core_DataAccess.GetAdodbRecordset(strSQL)



    If Core_VersionCheck.getVersion <> rst!ConfiguratorVersion Then

        MsgBox ("Please wait while the program updates"), vbCritical
    
        'Update CDI tables
        Call InstallDependencies

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

    Dim rst As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo ErrorHandler

    sUserName = GetCurrentUser()

    UserCheck = False


    strSQL = FormatString("SELECT * FROM tblConfiguratorUser WHERE [User] = '{0}' AND UserEmail IS NOT NULL", sUserName)

    Set rst = Core_DataAccess.GetAdodbRecordset(strSQL)


    If rst.EOF = True Then

        DoCmd.OpenForm "frmUserEmail"
    
        MsgBox "Please add your AFL Global email address for notifications", vbCritical
        
        Call InstallDependencies
    
        Exit Function

    End If

    UserCheck = True

    sUser = rst!UserTypeID
    iCutSheetApprover = rst!CutSheetApprover
    'UserCheck = rst!UserTypeID


    rst.Close

ErrorHandler:
    Exit Function


End Function

Sub InstallDependencies()

    Call RemoveOldConfig
    Call CreateConfigFolder
    Call MakeShortCut
    Call getBatFile
    Call UpdateCdiTables

End Sub

Public Function ItemDateCreated(Oracle As String) As Boolean

    Dim rst As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "SELECT cast([DateCreated] as date) as CreationDate, cast(getdate() as date) as CurrentDate "
    strSQL = strSQL & "FROM [Basic Product Construction] "
    strSQL = FormatString(" {0} WHERE [New Oracle Part #] = '{1}'", strSQL, Oracle)

    Set rst = Core_DataAccess.GetAdodbRecordset(strSQL)
    
    ItemDateCreated = (rst!CreationDate = rst!CurrentDate)

    rst.Close

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

Public Function BaseID_DateCreated(BaseID As Integer) As Boolean

    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim flag As Boolean
    

    strSQL = "SELECT cast([DateCreated] as date) as CreationDate, cast(getdate() as date) as CurrentDate "
    strSQL = strSQL & "FROM [tblCableConstructions] "
    strSQL = FormatString(" {0} WHERE [BaseID] = {1}", strSQL, BaseID)

    Set rst = Core_DataAccess.GetAdodbRecordset(strSQL)
    
    If rst.EOF Then Exit Function

    flag = (rst!CreationDate = rst!CurrentDate)

    BaseID_DateCreated = flag

    rst.Close

End Function

Public Sub MakeShortCut()
    '''Create shortcut of the Configurator program on the user's desktop'''
    Dim pathShortcut

    pathShortcut = getUserProfile & "\Desktop\" & ApplicationDirectory.getApplicationName & ".lnk"
    'If CommonFunctions.FileExists(pathShortcut) Then
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

Private Sub CreateConfigFolder()                 'Create Premise Configurator file and add the icon
    '''Create folder to house the Configurator program'''
    Dim path As String
    Dim localIconPath As String
    localIconPath = ApplicationDirectory.getUserDirectory & "PremiseConfiguratorIcon.ico"

    path = ApplicationDirectory.getUserDirectory

    If CommonFunctions.FileExists(path) = False Then

        MkDir ApplicationDirectory.getUserDirectory

    End If

    CreateIcon (localIconPath)

End Sub

Private Sub RemoveOldConfig()                    ' iterate through all possible extensions of the Premise Configurator on the users desktop
    '''Remove old versions of the Configurator program'''
    Dim oldFile() As String
    Dim File As String
    Dim oldConfig As String
    Dim fileExt As New Collection
    Dim Ext As Variant
    oldFile = Split(ApplicationDirectory.getApplicationName, ".")

    fileExt.Add ".mdb"
    fileExt.Add ".accdb"
    fileExt.Add ".accde"
    fileExt.Add "." & oldFile(1)

    File = oldFile(0)

    For Each Ext In fileExt

        oldConfig = getUserProfile & "\Desktop\" & File & Ext
    
        If CurrentProject.FullName <> oldConfig Then
            DeleteFile (oldConfig)
        End If

    Next

    subRefreshDesktop


End Sub

Private Sub CreateIcon(localIconPath As String)

    If CommonFunctions.FileExists(localIconPath) Then
        Exit Sub
    End If

    FileCopy ApplicationDirectory.getApplicationDirectory & "\PremiseConfiguratorIcon.ico", localIconPath

End Sub

Private Sub getBatFile()
    '''get the Configurator program .bat file and add it to user's desktop'''

    If CommonFunctions.FileExists(ApplicationDirectory.getClientApplicationBatFileDirectory) Then
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

    If CommonFunctions.FileExists(ClientPath) = False Then

        MkDir "C:\Applications"

    End If


    FileCopy RemotePath, ClientPath


End Sub

Public Sub UpdateUserLastLogin()

    Dim proc As String

    proc = "[Users].[usp_UpdateUserLastLogin]"
    
    Call Core_DataAccess.ExecuteSqlProcedureNoReturnOutput(proc, GetCurrentUser)


End Sub

Private Sub AccessVersionCheck()
    '''Check the version of the access the user is using'''
    If Application.Version < 15 Then
        MsgBox "Please update to MS Access 2013 or newer to ensure all features work properly.", vbCritical
    End If


End Sub



Public Sub ConnectionType()

Dim oObject
Dim adapter
Dim item

Set oObject = GetObject("WINMGMTS:\\.\ROOT\cimv2")

Set adapter = oObject.InstancesOf("Win32_NetworkAdapter")

    For Each item In adapter

        If item.NetconnectionID <> "null" Then
         Debug.Print item.NetconnectionID
         Debug.Print item.AdapterType
         Debug.Print item.NetConnectionStatus
     End If
    Next
    
End Sub
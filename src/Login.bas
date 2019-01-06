Attribute VB_Name = "Login"
Option Compare Database
Option Explicit
Private User As New clsUser
Private startUpForm As String

Sub OpenRoleBasedForm()
'''Determine the correct form to open based on user's role
Dim i As Variant
Dim bDefault As Boolean
Dim dict As New Scripting.Dictionary
Dim defaultResponsibility As Integer


defaultResponsibility = 7
bDefault = False

If Not User.Load Then
    MsgBox "Error logging on.  Please contact product engineering", vbCritical
    Exit Sub
End If

Set dict = User.getResponsibilities

If dict.count = 0 Then
    If checkResponsibility = False Then
        MsgBox "A log on error has occurred.  Please contact Product Engineering", vbCritical
        Exit Sub
    End If
End If

    

For Each i In dict.Keys
    If i = defaultResponsibility Then
        bDefault = True
        startUpForm = dict(i)
    End If
Next i

If bDefault = False Or dict.count > 1 Then
    startUpForm = "frmSwitchBoard"

End If


DoCmd.OpenForm startUpForm

Call NavigationPane

dict.RemoveAll




End Sub
Function checkResponsibility() As Boolean
'Check to ensure a responsibility has been set for the user
'if not then set the default responsibility

Dim cnn As New ADODB.Connection
Set cnn = New ADODB.Connection
Dim provstr As String
Dim sqlProcedure As String, sQry$
Dim rst As New ADODB.Recordset

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = cPremiseServerConnection
cnn.Open provstr

On Error GoTo ErrorHandler

checkResponsibility = False


sQry = "SELECT UserID FROM Users.vUserResponsibility WHERE [User] = '" & VersionCheck.GetCurrentUser & "'"

rst.Open sQry, cnn, adOpenKeyset, adLockOptimistic


If rst.EOF = True Then


sqlProcedure = "DECLARE @return_value INT  EXEC @return_value = [Users].[usp_ResponsibilityAddDefault]  SELECT  'Return Value' = @return_value"

    rst.Close
    rst.Open (sqlProcedure), cnn, adOpenKeyset, adLockOptimistic
    
    If rst![Return Value] = 0 Then
        checkResponsibility = True
    End If

Else

checkResponsibility = True
    

End If



rst.Close
cnn.Close

ErrorHandler:
Exit Function


End Function

Sub NavigationPane()
Dim i As Variant
Dim paneLock As Boolean

paneLock = False

For Each i In User.getResponsibilities
    If i = 5 Then
        paneLock = True
    End If

Next

DoCmd.LockNavigationPane Not paneLock
ChangeProperty "AllowSpecialKeys", dbBoolean, paneLock




End Sub


Sub testCodeDBProperties()
On Error Resume Next

Dim prop As Variant

For Each prop In CodeDb.Properties
'Debug.Print prop.Name & ": " & prop & "type: " & VarType(prop)
Next prop
End Sub

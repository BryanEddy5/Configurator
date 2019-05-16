Option Compare Database
Option Explicit
Private User As New clsUser
Private startUpForm As String


Sub OpenRoleBasedForm()
'''Determine the correct form to open based on user's role
Dim i As Variant
Dim bDefault As Boolean
Dim Dict As New Scripting.Dictionary
Dim defaultResponsibility As Integer


defaultResponsibility = 7
bDefault = False

If Not User.Load Then
    MsgBox "Error logging on.  Please contact product engineering", vbCritical
    Exit Sub
End If

Set Dict = User.getResponsibilities

If Dict.Count = 0 Then
    If checkResponsibility = False Then
        MsgBox "A log on error has occurred.  Please contact Product Engineering", vbCritical
        Exit Sub
    End If
End If

    

For Each i In Dict.Keys
    If i = defaultResponsibility Then
        bDefault = True
        startUpForm = Dict(i)
    End If
Next i

If bDefault = False Or Dict.Count > 1 Then
    startUpForm = "frmSwitchBoard"

End If

If User.UserId = 1 Then
    
    

End If


DoCmd.OpenForm startUpForm

Call NavigationPane

Dict.RemoveAll




End Sub
Function checkResponsibility() As Boolean
'Check to ensure a responsibility has been set for the user
'if not then set the default responsibility


Dim SqlProcedure As String, sQry$
Dim rst As New ADODB.Recordset



On Error GoTo ErrorHandler

checkResponsibility = False


sQry = "SELECT UserID FROM Users.vUserResponsibility WHERE [User] = '" & Core_VersionCheck.GetCurrentUser & "'"

Set rst = Core_DataAccess.GetAdodbRecordset(sQry)


If rst.EOF = True Then


SqlProcedure = "DECLARE @return_value INT  EXEC @return_value = [Users].[usp_ResponsibilityAddDefault]  SELECT  'Return Value' = @return_value"

   
    Set rst = Core_DataAccess.GetAdodbRecordset(SqlProcedure)
    
    If rst![Return Value] = 0 Then
        checkResponsibility = True
    End If

Else

checkResponsibility = True
    

End If

rst.Close


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
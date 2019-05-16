Option Explicit
Option Compare Database

Sub SendEmail(Email_Subject As String, Email_Send_To, Email_Body As String, _
                Optional Email_Cc As String, Optional Email_Bcc As String, Optional Email_Send_From As String)

Dim Mail_Object, Mail_Single As Variant
On Error GoTo debugs
Set Mail_Object = CreateObject("Outlook.Application")
Set Mail_Single = Mail_Object.CreateItem(0)
With Mail_Single
    .Subject = Email_Subject
    .To = Email_Send_To
    .cc = Email_Cc
    .BCC = Email_Bcc
    .Body = Email_Body
    .send
End With
'MsgBox "Email was successful"
debugs:
If Err.Description <> "" Then MsgBox Err.Description
End Sub

Public Sub CutSheetEmail(Base As String, item$)

Dim strSQL As String
Dim rst As New ADODB.Recordset
Dim sEmail_To As String
Dim sEmail_Body As String
Dim sEmail_Subject As String
Dim sEmail_Cc As String


strSQL = "SELECT UserEmail "
strSQL = strSQL & " FROM tblConfiguratorUser "
strSQL = strSQL & " WHERE CutSheetApprover = 1 OR CutSheetApprover = 2"

Set rst = Core_DataAccess.GetAdodbRecordset(strSQL)

Do Until rst.EOF
    sEmail_To = rst!userEmail & "; " & sEmail_To
    rst.MoveNext
Loop


sEmail_To = sEmail_To & ";" & getCurrentUserEmail
sEmail_Subject = "Premise Spec Sheet Request"
sEmail_Body = FormatString("A spec sheet approval has been requested for construction {0} and item {1}.", Base, item)
sEmail_Cc = "bryan.eddy@aflglobal.com"

Call SendEmail(sEmail_Subject, sEmail_To, sEmail_Body, , sEmail_Cc)


End Sub


Function CutSheetEmailApproved(Base As String, sEmail_To, item$)
'''Need to add requested item to email notification of approved constructions
Dim sEmail_Body As String
Dim sEmail_Subject As String
Dim sEmail_Cc As String


sEmail_Subject = "Premise Spec Sheet Approval - " & Base & "; " & item
sEmail_Body = "Approved: Premise spec sheets have been approved for Construction " & Base & " and item " & item & "."
sEmail_Cc = "bryan.eddy@aflglobal.com"

Call SendEmail(sEmail_Subject, sEmail_To, sEmail_Body, , sEmail_Cc)

End Function
Function getCurrentUserEmail()
Dim olNS As Outlook.NameSpace
Dim olFol As Outlook.Folder


Set olNS = Outlook.GetNamespace("MAPI")
Set olFol = olNS.GetDefaultFolder(olFolderInbox)

getCurrentUserEmail = olFol.Parent.name


End Function
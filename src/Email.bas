Attribute VB_Name = "Email"
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

Function CutSheetEmail(Base As String, Item$)

Dim strSQL As String
Dim rst As New ADODB.Recordset
Dim cnn As ADODB.Connection
Dim sEmail_To As String
Dim sEmail_Body As String
Dim sEmail_Subject As String
Dim sEmail_Cc As String



Set cnn = CurrentProject.Connection

strSQL = "SELECT UserEmail "
strSQL = strSQL & " FROM tblConfiguratorUser "
strSQL = strSQL & " WHERE CutSheetApprover = 1 OR CutSheetApprover = 2"

rst.Open (strSQL), cnn, adOpenStatic, adLockReadOnly

Do Until rst.EOF
    sEmail_To = rst!userEmail & "; " & sEmail_To
    rst.MoveNext
Loop


sEmail_To = sEmail_To & getCurrentUserEmail
sEmail_Subject = "Premise Spec Sheet Request"
sEmail_Body = "A spec sheet approval has been requested for construction " & Base & " and item " & Item & "."
sEmail_Cc = "bryan.eddy@aflglobal.com"

Call SendEmail(sEmail_Subject, sEmail_To, sEmail_Body, , sEmail_Cc)


End Function


Function CutSheetEmailApproved(Base As String, sEmail_To, Item$)
'''Need to add requested item to email notification of approved constructions
Dim sEmail_Body As String
Dim sEmail_Subject As String
Dim sEmail_Cc As String


sEmail_Subject = "Premise Spec Sheet Approval - " & Base & "; " & Item
sEmail_Body = "Approved: Premise spec sheets have been approved for Construction " & Base & " and item " & Item & "."
sEmail_Cc = "bryan.eddy@aflglobal.com"

Call SendEmail(sEmail_Subject, sEmail_To, sEmail_Body, , sEmail_Cc)

End Function
Function getCurrentUserEmail()
Dim olNS As Outlook.NameSpace
Dim olFol As Outlook.Folder
Dim userEmail$

Set olNS = Outlook.GetNamespace("MAPI")
Set olFol = olNS.GetDefaultFolder(olFolderInbox)

userEmail = olFol.Parent.Name


End Function

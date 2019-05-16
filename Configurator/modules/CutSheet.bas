Option Compare Database
Option Explicit
Dim SpecSheet As New clsCutSheet


''' Possibly no longer used '''
Function CutSheet_CheckApproval(sBase As String, ByRef sEmail As String) As Boolean
''Check to see if the cut sheet has been approved'''

If sBase = "" Then
Exit Function
End If

Dim strSQL As String
Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection


Set cnn = CurrentProject.connection

strSQL = "SELECT Commercial_Approval, Technical_Approval, UserEmail "
strSQL = strSQL & " FROM tblCutSheetApproval INNER JOIN tblConfiguratorUser ON tblCutSheetApproval.Requestor = tblConfiguratorUser.SQL_User "
strSQL = strSQL & " WHERE Base = '" & sBase & "'"

rst.Open (strSQL), cnn, adOpenStatic, adLockReadOnly

If rst.EOF Then
MsgBox "An error occurred attempting to notify the requestor of approval"

Exit Function

End If


If rst!Technical_Approval <> 0 And rst!Commercial_Approval <> 0 Then

CutSheet_CheckApproval = True
sEmail = rst!userEmail

'Debug.Print sEmail
End If


cnn.Close

End Function


Function BaseQueuedForApproval() As ADODB.Recordset


Dim strSQL As String
Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection


Set cnn = CurrentProject.connection

strSQL = "SELECT Base, Oracle "
strSQL = strSQL & " FROM tblCutSheetApproval_Local WHERE Technical_Approval <> 0"


rst.Open (strSQL), cnn, adOpenStatic, adLockReadOnly

Set BaseQueuedForApproval = rst

'Do Until rst.EOF
'    BaseQueuedForApproval = rst!Base & ";" & BaseQueuedForApproval
'    rst.MoveNext
'Loop


End Function

Sub RequestCutSheet(OracleItem As String, ConstuctionBase As String)
'''Method updates the approval status of base(s) and fires off emails notifying the appropriate_ '''
'''parties that an cut sheet approval has been sent'''
Dim cutSheet As New clsCutSheet
Dim CoreItem As New clsCutSheet
Dim item As String, Base As String


cutSheet.Load (OracleItem)


If cutSheet.getCutSheetApproval Or cutSheet.Requested Then
    Exit Sub
End If


On Error GoTo ErrorHandler:

DoCmd.SetWarnings False


If OracleItem Like "*IA" Then


    item = FindCoreCable(OracleItem)
    CoreItem.Load (item)
    Base = CoreItem.Base
    If CoreItem.Requested = False Then 'If core cable hasn't been requested, then send notification and request for core cable
        Call CutSheetApproval(Base, 3, item)
        Call CutSheetEmail(Base, item)

    End If

End If



    Call CutSheetApproval(cutSheet.Base, 3, OracleItem)
    Call CutSheetEmail(cutSheet.Base, OracleItem)
    


ErrorExit:
Exit Sub

ErrorHandler:
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit

End Sub


Sub OpenCutSheet(item As String)
'''Sub is to open the report of the cut sheet'''

If SpecSheet.Load(item) = 0 Then
    MsgBox "No spec sheet is available for this item." & vbCrLf & "Contact Product Engineering if spec sheet is required.", vbCritical
    Exit Sub
End If

DoCmd.OpenReport "rptCutSheet", acViewPreview, , "Oracle ='" & item & "'", acWindowNormal

End Sub



Function CutSheetApproval(Base As String, ApprovalProcedure As Integer, Optional item As String)


'''Function is to call the procedure to approve base(s) from an authorized approver'''
Dim strStoredProcedure As String
Dim rst As New ADODB.Recordset



On Error GoTo ErrorHandler:

strStoredProcedure = "usp_CutSheet_Approval"

Call Core_DataAccess.ExecuteSqlProcedureNoReturnOutput(strStoredProcedure, Base, ApprovalProcedure, item)


ErrorExit:
Exit Function

ErrorHandler:
MsgBox "An error occurred during the approval submission.", vbCritical
Resume ErrorExit


End Function

Public Function getBaseImage(Base As String, Color As String, ImageTypeID As Integer, Armored As Boolean) As String
'''Function gets the image for the cable Cut Sheet'''

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim rst As New ADODB.Recordset
Dim provstr As String
Dim sql As String
Dim Armor As Integer

Armor = CInt(Abs(Armored))



' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr



sql = "SELECT ImageFileType,ImageID FROM CutSheet.vImage where Base = '" & Base & "' AND Color = '" & Color & "' AND ImageTypeID = " & ImageTypeID & " AND Armored = " & Armor

rst.Open (sql), cnn, adOpenKeyset, adLockOptimistic


If rst.EOF Then
getBaseImage = ""
Exit Function
End If


    With rst
        getBaseImage = Nz(!ImageID) & Nz(!ImageFileType)
    End With
    
    



rst.Close
cnn.Close



'
End Function

Public Function MechPropertiesFormat(UOM As String, NumDecimalPlaces As Integer, Measurment As Double) As String

Dim sFormat As String
Dim i As Integer

sFormat = "#,##."

For i = 0 To NumDecimalPlaces - 1

sFormat = sFormat & "0"

Next

MechPropertiesFormat = Format(Round(Measurment, NumDecimalPlaces), sFormat) & " " & UOM



End Function

Public Function AddImageToBase(Bases As DAO.Recordset, ImageGroupID%) As Boolean

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sql As String
Dim rst As New ADODB.Recordset



On Error GoTo ErrorHandler:


AddImageToBase = False
    cnn.Provider = "sqloledb"
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr


Do Until Bases.EOF

    
    sql = "SELECT ImageGroupID FROM tblCableConstructionReferences WHERE SetupID = " & Bases!SetupID

    
    rst.Open (sql), cnn, adOpenKeyset, adLockOptimistic
    
    With rst
        !ImageGroupID = ImageGroupID
        .Update
    End With
    Bases.Delete
    
    
    Bases.MoveNext
    
    rst.Close

Loop


AddImageToBase = True



ErrorExit:
Exit Function

ErrorHandler:
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit

End Function

Public Function FiberCode(Fiber As String, CorningSpecific As Boolean) As String

Dim strSQL As String
Dim rst As New ADODB.Recordset
Dim cnn As ADODB.connection

Set cnn = CurrentProject.connection

strSQL = "SELECT CodeFiberType, CorningSpecific "
strSQL = strSQL & " FROM tbl_Fibers INNER JOIN dbo_tblCutSheetOpticalFiberChart ON tbl_Fibers.OpticalID = dbo_tblCutSheetOpticalFiberChart.OpticalID "
strSQL = strSQL & " WHERE Fiber  = '" & Fiber & "'"

rst.Open (strSQL), cnn, adOpenStatic, adLockReadOnly

If CorningSpecific = True Then
    FiberCode = rst!CorningSpecific
Else
    FiberCode = rst!CodeFiberType
End If



End Function
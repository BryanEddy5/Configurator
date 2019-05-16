Option Explicit
Option Compare Database

Global clnClient As New Collection     'Instances of frmDoublePassCable
Public Enum InstanceFormType
    ItemOracle
    ItemID
    ConstSetupID
    ConstBase
    FindItem
    UnitSeriesfrm
End Enum
    
Function OpenAClient(ItemType As InstanceFormType, FilterFor As String, Optional Location As Integer)
    'Purpose:    Open an independent instance of form frmClient.
    Dim frm As Form
    Dim strFilter$
    Location = Nz(Location, 2)
    
    Select Case ItemType
        Case ItemOracle                         'Shows form of discrete items created
            If Location = 1 Then
                Set frm = New Form_frmSinglePassCable
            
            Else
                Set frm = New Form_frmDoublePassCable
            End If
    
            strFilter = "Oracle = '"
            FilterFor = FilterFor & "'"
            
        Case ItemID                       'Shows form of discrete items created
            If Location = 1 Then
                Set frm = New Form_frmSinglePassCable
            
            Else
                Set frm = New Form_frmDoublePassCable
            End If
    
            strFilter = "ItemID = "
    
        Case ConstSetupID                   'Displays form of constructions
    
            If Location = 1 Then
                Set frm = New Form_frmConstructionSinglePass
            
            ElseIf Location = 2 Then
                Set frm = New Form_frmConstructionDoublePass
            ElseIf Location = 0 Then
                MsgBox "No construction data found", vbCritical
                Set frm = New Form_frmConstructionSinglePass
                'Exit Function
            End If
            strFilter = "SetupID = "

        Case ConstBase                    'Displays form of constructions
    
            If Location = 1 Then
                Set frm = New Form_frmConstructionSinglePass
            
            ElseIf Location = 2 Then
                Set frm = New Form_frmConstructionDoublePass
            ElseIf Location = 0 Then
                MsgBox "No construction data found", vbCritical
                Set frm = New Form_frmConstructionSinglePass
                'Exit Function
            End If
            strFilter = "Base = '"
            FilterFor = FilterFor & "'"
        Case FindItem
            Set frm = New Form_frmFindItem
        Case UnitSeriesfrm
            Set frm = New Form_frmFindUnitSeries
            
    
    End Select
    'Open a new instance, show it, and set a caption.
    frm.Visible = True
    frm.Caption = frm.hWnd & ", opened " & Now()
    frm.Filter = strFilter & FilterFor
    frm.FilterOn = True
    'Append it to our collection.
    clnClient.Add item:=frm, Key:=CStr(frm.hWnd)

    'Set frm = Nothing
End Function

Function CloseAllClients() As Boolean
    'Purpose: Close all instances in the clnClient collection.
    'Note: Leaves the copy opened directly from database window/nav pane.
    Dim lngKt As Long
    Dim lngI As Long

    lngKt = clnClient.Count
    For lngI = 1 To lngKt
        clnClient.Remove 1
        CloseAllClients = True
    Next
End Function

Function CloseClient(strHwnd As String) As Boolean

    'Purpose: Remove this instance from clnClient collection.
    On Error GoTo ErrorHandler
    
    Dim obj As Object        'Object in clnClient

    Dim blnRemove As Boolean  'Flag to remove it.

    'Check if this instance is in the collection.
    For Each obj In clnClient
        If obj.hWnd = strHwnd Then
            blnRemove = True
            Exit For
        End If
    Next

    'Deassign the object and remove from collection.
    Set obj = Nothing
    If blnRemove Then
        clnClient.Remove CStr(strHwnd)
        CloseClient = True
    End If
    
ErrorExit:
    Exit Function

ErrorHandler:
    'DoCmd.Close acForm, Me.Name
    Resume ErrorExit:
End Function

Function OpenAClientForm(ItemType As Integer, FilterFor As String, ReportName As String)
    'Purpose:    Open an independent instance of form frmClient.
    Dim rpt As Report, strFilter$
    
    
    Select Case ItemType
        Case 0
    'Shows form of discrete items created
        Set rpt = New Report_rptCutSheet
    
        strFilter = "Oracle = '"
        Case 1
        
        
    
        Case 3        'Displays form of constructions
            Set rpt = New Form_frmConstructionSinglePass
        strFilter = "Base = '"

    End Select
    
        'Open a new instance, show it, and set a caption.
    rpt.Visible = True
    'rpt.Caption = frm.Hwnd & ", opened " & Now()
    rpt.Filter = strFilter & FilterFor & "'"
    rpt.FilterOn = True
    'rpt.Requery
    'Append it to our collection.
    clnClient.Add item:=rpt, Key:=CStr(rpt.hWnd)
    
    'Call Report_rptCutSheet.Initialize

    'Set frm = Nothing
End Function
Option Explicit
Option Compare Database

'------------------------------------------------------------
' mcrItemInterfaceDisplay
'
'------------------------------------------------------------
Sub sItemInterfaceDisplay()
On Error GoTo mcrItemInterfaceDisplay_Err

    If (Forms!frmItemInterface!OrganizationCode = "000") Then
        ' frmItemInterfaceCatalog - Yes
        Forms!frmItemInterface!frmItemInterfaceCatalog.Form.Visible = True

        ' frmItemInterfaceCategory - Yes
        Forms!frmItemInterface!frmItemInterfaceCategory.Form.Visible = True

        ' frmItemInterfaceBOM - No
        Forms!frmItemInterface!frmItemInterfaceBOM.Form.Visible = False

        ' frmItemInterfaceRouting - No
        Forms!frmItemInterface!frmItemInterfaceRouting.Form.Visible = False
    End If
    If (Forms!frmItemInterface!OrganizationCode <> "000") Then
        ' frmItemInterfaceCatalog - No
        Forms!frmItemInterface!frmItemInterfaceCatalog.Form.Visible = False

        ' frmItemInterfaceCategory - yes
        Forms!frmItemInterface!frmItemInterfaceCategory.Form.Visible = False

        ' frmItemInterfaceBOM - Yes
        Forms!frmItemInterface!frmItemInterfaceBOM.Form.Visible = True

        ' frmItemInterfaceRouting - Yes
        Forms!frmItemInterface!frmItemInterfaceRouting.Form.Visible = True
            End If


mcrItemInterfaceDisplay_Exit:
    Exit Sub

mcrItemInterfaceDisplay_Err:
    MsgBox Error$
    Resume mcrItemInterfaceDisplay_Exit

End Sub
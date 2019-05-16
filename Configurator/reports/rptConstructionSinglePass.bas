Version =20
VersionRequired =20
Begin Report
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9240
    DatasheetFontHeight =10
    ItemSuffix =43
    Left =4680
    Top =885
    DatasheetGridlinesColor =12632256
    Filter ="base ='SE000-2.0A'"
    RecSrcDt = Begin
        0xab469c8bcefee440
    End
    RecordSource ="SELECT tblCableConstructionReferences.Base, tblCableConstructionReferences.DateC"
        "reated, tblCableConstructionReferences.CreatedBy, tblCableConstructionReferences"
        ".RevisionDate, tblCableConstructionReferences.RevisedBy, tblCableConstructionRef"
        "erences.DesignCode, tblCableConstructionReferences.DesignExtension, tblCableCons"
        "tructionReferences.ConstructionDescription, tblCableConstructionReferences.BaseI"
        "D, tblCableConstructionReferences.NumSubFillers, tblCableConstructionReferences."
        "Active, tblCableConstructionReferences.PrefixID, tblCableConstructionReferences."
        "ProductID, tblCableConstructionReferences.SetupID, tblCableConstructionReference"
        "s.FamilyID, tblCableConstructionReferences.UnitIDTypeNumber, tblCableConstructio"
        "nReferences.InactiveReason, tblCableConstructionReferences.InactiveReason, tblCa"
        "bleConstructionReferences.TBType FROM tblCableConstructionReferences; "
    Caption ="Cable Construction"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a00500000000000018240000783c000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    FitToPage =1
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =3
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =90.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =15480
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    Left =60
                    Top =60
                    Width =9180
                    Height =15420
                    BorderColor =15916467
                    Name ="Child41"
                    SourceObject ="Form.frmConstructionSinglePass"
                    LinkChildFields ="Base"
                    LinkMasterFields ="Base"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =15480
                    Begin
                        Begin Label
                            Left =60
                            Width =660
                            Height =240
                            Name ="Label42"
                            Caption ="Child41:"
                            LayoutCachedLeft =60
                            LayoutCachedWidth =720
                            LayoutCachedHeight =240
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database




Private Sub btnBurnListing_Click()
DoCmd.OpenForm "frmFindItemBurnListing"
End Sub

Private Sub btnClose_Click()
    'Purpose: Remove this instance from clnClient collection.
    On Error GoTo ErrorHandler
    
    Dim obj As Object

    For Each obj In clnClient
        If obj.hWnd = Me.hWnd Then
            clnClient.Remove CStr(Me.hWnd)
        End If
    Next
    
ErrorExit:
    Exit Sub
    
ErrorHandler:

    Resume ErrorExit:
End Sub


Public Sub btnEdit_Click()

UpdateItem (False)
End Sub

Public Sub UpdateItem(UpdateButton As Boolean)
Dim ctl As control

If sUser = 1 Or UpdateButton = True Then
    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
            If ctl.Locked <> False Then
                ctl.Locked = False
            End If
        End If
    Next ctl
    For Each ctl In Me!frmCableConstruction.Form.Controls
            If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
            If ctl.Locked <> False Then
                ctl.Locked = False
            End If
        End If
    Next ctl
        For Each ctl In Me!frmCableConstructionPrefixID.Form.Controls
            If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
            If ctl.Locked <> False Then
                ctl.Locked = False
            End If
        End If
    Next ctl
End If
'Me.Revision_History.Locked = True
'Me.Revision_Date.Locked = True
'Me.BaseID.Locked = True



End Sub

Private Sub btnMils_Click()
On Error GoTo Error_Handler

DoCmd.OpenForm "frmMilsConversion"

Forms!frmMilsConversion!NominalWall = Round(CDbl(Me!frmCableConstruction.Form!NominalWall.NominalWall) * 39.37, 1)
Forms!frmMilsConversion!MinAveWall = Round(CDbl(Me!frmCableConstruction.Form!NominalWall.MinAveWall) * 39.37, 1)
Forms!frmMilsConversion!MinSpotWall = Round(CDbl(Me!frmCableConstruction.Form!NominalWall.MinSpotWall) * 39.37, 1)
Forms!frmMilsConversion!MaxAveWall = Round(CDbl(Me!frmCableConstruction.Form!NominalWall.MaxAveWall) * 39.37, 1)

Error_Exit:
Exit Sub

Error_Handler:
'MsgBox Err.Description & " " & Err.Number
Resume Error_Exit
End Sub





Private Sub btnPrintPreview_Click()

'DoCmd.OpenReport "rptSinglePassCable", acViewPreview, , "Oracle ='" & Me.Item & "'", acWindowNormal
'DoCmd.RunCommand acCmdPreviewTwoPages

End Sub

Private Sub btnRemoveFilter_Click()
Me.Requery
End Sub

Private Sub btnUpdate_Click()
DoCmd.OpenForm "frmRevisionReason"
Forms!frmRevisionReason!strHwnd = Me.Caption 'CStr(Me.Hwnd)

'Me.Revision_Date = Date
'Me.Revision_Letter = Int(Me.Revision_Letter) + Int(1)


End Sub


Private Sub Form_Load()
Me.move 0, 0

Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
        If ctl.Locked <> True Then
            ctl.Locked = True
        End If
    End If
Next ctl
'Me.Reason.Locked = False
'Me.Active.Locked = False
End Sub

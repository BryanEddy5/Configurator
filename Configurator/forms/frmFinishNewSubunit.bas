Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =85
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =26580
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =405
    Top =1065
    Right =22980
    Bottom =13455
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe82a8e49398de440
    End
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnGotFocus ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =13440
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =87
                    SpecialEffect =5
                    BorderWidth =3
                    Left =120
                    Top =4020
                    Width =25740
                    Height =4800
                    Name ="frmUnitSeriesSubunits"
                    SourceObject ="Form.frmUnitSeriesSubunits"

                    LayoutCachedLeft =120
                    LayoutCachedTop =4020
                    LayoutCachedWidth =25860
                    LayoutCachedHeight =8820
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =3780
                            Width =1650
                            Height =240
                            Name ="frmUnitSeriesSubunits Label"
                            Caption ="Unit Series"
                            EventProcPrefix ="frmUnitSeriesSubunits_Label"
                            LayoutCachedLeft =120
                            LayoutCachedTop =3780
                            LayoutCachedWidth =1770
                            LayoutCachedHeight =4020
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =180
                    Top =9060
                    Width =25800
                    Height =3720
                    TabIndex =1
                    Name ="frmNewSubUnit"
                    SourceObject ="Form.frmNewSubUnit"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =8820
                            Width =1185
                            Height =240
                            Name ="frmNewSubUnit Label"
                            Caption ="New Subunit"
                            EventProcPrefix ="frmNewSubUnit_Label"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =240
                    Top =600
                    Width =25680
                    Height =3120
                    TabIndex =2
                    Name ="frmFinalCableReview"
                    SourceObject ="Form.frmFinalCableReview"

                    LayoutCachedLeft =240
                    LayoutCachedTop =600
                    LayoutCachedWidth =25920
                    LayoutCachedHeight =3720
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =300
                            Top =360
                            Width =1575
                            Height =240
                            Name ="frmFinalCableReview Label"
                            Caption ="Basic (Cable/Sheath)"
                            EventProcPrefix ="frmFinalCableReview_Label"
                            LayoutCachedLeft =300
                            LayoutCachedTop =360
                            LayoutCachedWidth =1875
                            LayoutCachedHeight =600
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1860
                    Top =12900
                    Width =1500
                    Height =480
                    FontWeight =700
                    TabIndex =3
                    ForeColor =0
                    Name ="btnLoad"
                    Caption ="Load All"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =300
                    Top =12900
                    Height =480
                    FontWeight =700
                    TabIndex =4
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub btnClose_Click()

    DoCmd.Close acForm, Me.name


End Sub

Private Sub btnLoad_Click()

    On Error GoTo Error_Handler
    DoCmd.SetWarnings False
    Dim control

    For Each control In Me.Controls
        If control.ControlType = acSubform Then
            control.Requery
        End If
    Next



    DoCmd.OpenQuery "qryLOADBasicProductConstruction"
    DoCmd.OpenQuery "qryLOADPremisUnitSeries"
    DoCmd.OpenQuery "qryLOADBasicCableHold_New"
    Me.btnClose.SetFocus
    Me.btnLoad.Enabled = False
    DoCmd.SetWarnings True



    Call CreateNewItem.InsertRequestToLog

    Beep
    MsgBox "Load complete"

    Call Form_Close

    Call btnClose_Click

Error_Exit:
    Exit Sub

Error_Handler:
    DoCmd.SetWarnings False
    MsgBox (Err.Description & " " & Err.Number)
    Resume Error_Exit



End Sub

Private Sub Form_Close()
    If CurrentProject.AllForms("frmCreateItem").IsLoaded Then
        Call Forms("frmCreateItem").btnClose_Click
    End If
End Sub

Private Sub Form_GotFocus()
    Me.Refresh

End Sub

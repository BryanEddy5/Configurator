Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14880
    DatasheetFontHeight =10
    ItemSuffix =118
    Left =12015
    Top =4365
    Right =27150
    Bottom =10290
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa085a14c6b17e540
    End
    RecordSource ="qryPremiseTBReplaceTB"
    Caption ="TB Replacement Options"
    OnClose ="mcrPremiseTBReplaceClose"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =1140
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =480
                    Top =780
                    Width =2940
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Existing TB Item"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =8040
                    Top =180
                    Width =900
                    FontSize =10
                    FontWeight =700
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="mcrClose"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =8100
                    Top =780
                    Width =1140
                    Height =300
                    FontSize =10
                    Name ="Label99"
                    Caption ="New TB Item"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =8100
                    LayoutCachedTop =780
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =1080
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =480
                    Top =360
                    Width =2940
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cboBaseSelect"
                    RowSourceType ="Table/Query"
                    RowSource ="qryPremiseTBReplaceBaseTB"
                    AfterUpdate ="mcrRequery"
                    FontName ="Times New Roman"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =7020
                    Top =180
                    Width =900
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdUpdate"
                    Caption ="Update"
                    OnClick ="mcrPremiseTBReplace"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =3600
                    Top =780
                    Width =1620
                    Height =300
                    FontSize =10
                    Name ="Label109"
                    Caption ="Catalog No"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3600
                    LayoutCachedTop =780
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =10260
                    Top =780
                    Width =1620
                    Height =300
                    FontSize =10
                    Name ="Label110"
                    Caption ="Catalog No"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =10260
                    LayoutCachedTop =780
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =1080
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =11160
                    Top =60
                    Width =2940
                    Height =300
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =3
                    Name ="FiberSelection"
                    RowSourceType ="Value List"
                    RowSource ="Fiber1;Fiber2;Fiber3"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =9960
                    Top =60
                    Width =1140
                    Height =300
                    FontSize =10
                    Name ="Label112"
                    Caption ="Fiber Select"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =11160
                    Top =420
                    Width =2940
                    Height =300
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =4
                    Name ="FiberType"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =480
                    Top =60
                    Width =2940
                    Height =300
                    ColumnWidth =465
                    FontSize =10
                    Name ="ComponentItemNumber"
                    ControlSource ="ComponentItemNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8160
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="NewComponentItem"
                    ControlSource ="NewComponentItem"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8160
                    LayoutCachedTop =60
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3600
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Text114"
                    ControlSource ="CopiedItemNo"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =10260
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="Text115"
                    ControlSource ="NewItemNo"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =10260
                    LayoutCachedTop =60
                    LayoutCachedWidth =12180
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =120
            BackColor =-2147483633
            Name ="FormFooter"
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

Private Sub FiberSelection_AfterUpdate()

    On Error GoTo Error_Handler

    Select Case Me.FiberSelection
    Case "Fiber1"
        Me.FiberType = [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![FiberType1]
    Case "Fiber2"
        Me.FiberType = [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![FiberType2]
    Case "Fiber3"
        Me.FiberType = [Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![FiberType3]
    End Select

    Me.Requery

    Me.Refresh
    Me.cboBaseSelect = Null
    Me.cboBaseSelect.Requery
    Me.cboBaseSelect = Me.cboBaseSelect.ItemData(0)

    DoCmd.RunMacro "mcrRequery"

Error_Exit:
    Exit Sub

Error_Handler:
    MsgBox Err.Description & " " & Err.Number
    Resume Error_Exit

End Sub

Private Sub Form_Load()




    Me.FiberSelection = Me.FiberSelection.ItemData(0)


    If Forms!frmItemInterface!frmItemInterfaceBOM.Form!ComponentItemNumber Like "pt*" Then
        Me.RecordSource = "qryPremiseTBReplaceTB"
        Me.cboBaseSelect.rowSource = "qryPremiseTBReplaceBaseTB"
    ElseIf Forms!frmItemInterface!frmItemInterfaceBOM.Form!ComponentItemNumber Like "rbn*" Or Forms!frmItemInterface!frmItemInterfaceBOM.Form!ComponentItemNumber Like "swr*" Then
        Me.RecordSource = "qryPremiseTBReplaceRibbon"
        Me.cboBaseSelect.rowSource = "qryPremiseTBReplaceBaseRibbon"
    End If

    Call FiberSelection_AfterUpdate



End Sub

Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =23
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =28800
    DatasheetFontHeight =10
    ItemSuffix =29
    Left =5235
    Top =6210
    Right =24285
    Bottom =13290
    DatasheetGridlinesColor =12632256
    Filter ="(((qryPrintLine.Cust Like \"BELDEN*\"))) And (qryPrintLine.PrintLine1 Like \"*DU"
        "PLEX*\")"
    OrderBy ="[qryPrintLine].[PrintLine2], [qryPrintLine].[PrintLine1], [qryPrintLine].[Item N"
        "o]"
    RecSrcDt = Begin
        0x671a4ef6b5a3e440
    End
    RecordSource ="qryPrintLine"
    Caption ="qryPrintLine"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
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
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =3600
                    Height =240
                    FontSize =9
                    FontWeight =600
                    Name ="Cust_Label"
                    Caption ="Customer"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3780
                    Top =60
                    Width =5040
                    Height =240
                    FontSize =9
                    FontWeight =600
                    Name ="PrintLine1_Label"
                    Caption ="Print Line 1"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8940
                    Top =60
                    Width =6900
                    Height =240
                    FontSize =9
                    FontWeight =600
                    Name ="PrintLine2_Label"
                    Caption ="Print Line 2"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =15960
                    Top =60
                    Width =6180
                    Height =240
                    FontSize =9
                    FontWeight =600
                    Name ="PrintLine3_Label"
                    Caption ="Print Line 3"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =22320
                    Top =60
                    Width =3420
                    Height =240
                    FontSize =9
                    FontWeight =600
                    Name ="Label13"
                    Caption ="Design Code"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =25920
                    Top =60
                    Width =720
                    Height =240
                    FontSize =9
                    FontWeight =600
                    Name ="Label15"
                    Caption ="UOM "
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =3600
                    Height =300
                    ColumnWidth =1650
                    FontSize =9
                    FontWeight =500
                    Name ="Cust"
                    ControlSource ="Cust"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3780
                    Top =60
                    Width =5040
                    Height =300
                    ColumnWidth =3000
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="PrintLine1"
                    ControlSource ="PrintLine1"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8940
                    Top =60
                    Width =6900
                    Height =300
                    ColumnWidth =3000
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="PrintLine2"
                    ControlSource ="PrintLine2"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15960
                    Top =60
                    Width =6180
                    Height =300
                    ColumnWidth =5940
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="PrintLine3"
                    ControlSource ="PrintLine3"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22320
                    Top =60
                    Width =3420
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="Text14"
                    ControlSource ="Item No"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25920
                    Top =60
                    Width =720
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="Text16"
                    ControlSource ="Print4"
                    FontName ="Tahoma"

                End
            End
        End
        Begin FormFooter
            Height =540
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin ToggleButton
                    OverlapFlags =85
                    Left =480
                    Top =60
                    Width =1080
                    Height =420
                    Name ="Toggle8"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =480
                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =60
                    Width =1380
                    Height =420
                    TabIndex =1
                    ForeColor =-2147483630
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =480
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
Option Compare Database

Private Sub btnRemoveFilter_Click()
    Me.Filter = ""
    Me.FilterOn = False
    Me.Requery
    Me.Refresh
End Sub

Private Sub Toggle8_Click()
    DoCmd.Close


End Sub

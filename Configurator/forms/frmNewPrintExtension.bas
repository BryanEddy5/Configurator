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
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6060
    DatasheetFontHeight =10
    ItemSuffix =65
    Left =11220
    Top =4440
    Right =17280
    Bottom =8565
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x32be2ee6d473e440
    End
    Caption ="Add A Print Extension"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =1620
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1380
                    Top =1020
                    Width =990
                    Height =450
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdYes"
                    Caption ="YES"
                    OnClick ="mcrNewPrintExtensionFromBaseYES"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2880
                    Top =1020
                    Width =990
                    Height =450
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdNo"
                    Caption ="NO"
                    OnClick ="mcrClose"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =180
                    Width =5580
                    Height =780
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    BackColor =-2147483633
                    ForeColor =0
                    Name ="Label58"
                    ControlSource ="=\"Have you successfully completed the loading of the base item:  \" & Forms!frm"
                        "MainMenu!NewItemNumber & \"?\""
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =60
            BackColor =14211288
            Name ="Detail"
        End
        Begin FormFooter
            Visible = NotDefault
            Height =2640
            BackColor =14211288
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1440
                    Top =2040
                    Width =990
                    Height =450
                    FontSize =10
                    FontWeight =700
                    Name ="cmdNewPrintExtension"
                    Caption ="Continue"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2880
                    Top =2040
                    Width =990
                    Height =450
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdCancel"
                    Caption ="Cancel"
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
                    Left =420
                    Top =180
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label40"
                    Caption ="Base Item"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2460
                    Top =180
                    Width =1740
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="ExistingBaseItem"
                    ControlSource ="=Forms!frmMainMenu!NewItemNumber"
                    Format =">"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =420
                    Top =600
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label56"
                    Caption ="New Print Extension"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4560
                    Top =180
                    Width =300
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =0
                    Name ="BaseItem"
                    ControlSource ="=Left(Forms!frmMainMenu!NewItemNumber,InStr(Forms!frmMainMenu!NewItemNumber,\"-\""
                        ")-1)"
                    Format =">"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =420
                    Top =1620
                    Width =1740
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label63"
                    Caption ="New Item Number"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2460
                    Top =1620
                    Width =1740
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="NewItemNumber"
                    ControlSource ="=IIf([NewPrintExtension] Is Null,\"\",[NewPrintExtension])"
                    Format =">"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1080
                    Top =1020
                    Width =4140
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label61"
                    Caption ="Enter ONLY the print extension (01, 02, 10, etc)"
                    FontName ="Times New Roman"
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =2460
                    Top =600
                    Width =1740
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    ForeColor =0
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="NewPrintExtension"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryPremiseItems.OracleItemNumber, qryPremiseItems.CatalogNumber, qryPremi"
                        "seItems.Customer, Left([OracleItemNumber],7) AS BaseItem, Mid([OracleItemNumber]"
                        ",8,5) AS Extension FROM qryPremiseItems WHERE (((Left([OracleItemNumber],7))=Lef"
                        "t([Forms]![frmMainMenu]![NewBaseItem],7)) AND ((Mid([OracleItemNumber],8,5)) Not"
                        " Like \"*-00\" And (Mid([OracleItemNumber],8,5)) Not Like \"*bx*\" And (Mid([Ora"
                        "cleItemNumber],8,5)) Not Like \"*ia*\"));"
                    ColumnWidths ="1440;2160;2880"
                    AfterUpdate ="mcrNewPrintExtensionSetMain"
                    FontName ="Times New Roman"
                    Format =">"

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

Private Sub cmdNewPrintExtension_Click()
    Call NewPrintExtensionFromBaseContinue
End Sub

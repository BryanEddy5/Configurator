Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7680
    DatasheetFontHeight =10
    ItemSuffix =23
    Left =7785
    Top =2775
    Right =15345
    Bottom =9570
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x26ac52989699e240
    End
    RecordSource ="qryItemCatalogInterface"
    Caption ="frmItemInterface"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
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
            Height =720
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =240
                    Top =420
                    Width =1080
                    Height =300
                    FontSize =10
                    Name ="Label15"
                    Caption ="Item Number"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2040
                    Top =420
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label16"
                    Caption ="Description"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =4500
                    Top =420
                    Width =2340
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Value"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =120
                    Top =120
                    Width =1320
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label24"
                    Caption ="Item Catalogs"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =5940
                    Top =60
                    Width =900
                    FontSize =10
                    FontWeight =700
                    Name ="cmdEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =330
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =240
                    Top =15
                    Width =1740
                    Height =300
                    ColumnWidth =1185
                    FontSize =10
                    Name ="ItemNumber"
                    ControlSource ="ItemNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =2040
                    Top =15
                    Width =2400
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =1
                    Name ="ElementName"
                    ControlSource ="ElementName"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =4500
                    Top =15
                    Width =2280
                    Height =300
                    ColumnWidth =465
                    FontSize =10
                    TabIndex =2
                    Name ="ElementValue"
                    ControlSource ="ElementValue"
                    FontName ="Times New Roman"

                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =6840
                    Top =75
                    TabIndex =3
                    Name ="AllowOverride"
                    ControlSource ="AllowOverride"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7140
                    Top =15
                    Width =366
                    Height =306
                    TabIndex =4
                    Name ="cmdDelete"
                    Caption ="Command25"
                    OnClick ="mcrDeleteRecord"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =900
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =180
                    Top =60
                    Width =6960
                    Height =780
                    Name ="frmItemInterfacePrint"
                    SourceObject ="Form.frmItemInterfacePrint"

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

Private Sub cmdEdit_Click()
DoCmd.OpenForm "frmPrintExtItemInterfaceCatalog"
End Sub

Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9300
    DatasheetFontHeight =10
    ItemSuffix =55
    Left =735
    Top =3075
    Right =10110
    Bottom =9510
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2e86c0d68f74e440
    End
    RecordSource ="qryItemInterfaceBOMView"
    Caption ="frmItemInterface"
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
            Height =1140
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Top =780
                    Width =480
                    Height =300
                    FontSize =10
                    Name ="Label15"
                    Caption ="Org"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =60
                    LayoutCachedTop =780
                    LayoutCachedWidth =540
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =600
                    Top =780
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label16"
                    Caption ="Item"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =600
                    LayoutCachedTop =780
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =2640
                    Top =540
                    Width =1380
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Component"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =540
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =4680
                    Top =540
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Label23"
                    Caption ="Oper"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =4680
                    LayoutCachedTop =540
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =5280
                    Top =540
                    Width =900
                    Height =300
                    FontSize =10
                    Name ="Label25"
                    Caption ="Effectivity"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =5280
                    LayoutCachedTop =540
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6240
                    Top =780
                    Width =900
                    Height =300
                    FontSize =10
                    Name ="Label27"
                    Caption ="Quantity"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =6240
                    LayoutCachedTop =780
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7200
                    Top =780
                    Width =300
                    Height =300
                    FontSize =10
                    Name ="Label29"
                    Caption ="Ct"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =7200
                    LayoutCachedTop =780
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =4680
                    Top =780
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Label31"
                    Caption ="Seq"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =4680
                    LayoutCachedTop =780
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =5280
                    Top =780
                    Width =900
                    Height =300
                    FontSize =10
                    Name ="Label32"
                    Caption ="Date"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =5280
                    LayoutCachedTop =780
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =60
                    Top =480
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label28"
                    Caption ="Bill of Materials"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =780
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =4080
                    Top =540
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Label34"
                    Caption ="Item"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =540
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =4080
                    Top =780
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Label35"
                    Caption ="Seq"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =780
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =8040
                    Top =600
                    Width =420
                    Height =300
                    FontSize =10
                    Name ="Label44"
                    Caption ="Layer"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =8040
                    LayoutCachedTop =600
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =900
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =8520
                    Top =600
                    Width =420
                    Height =300
                    FontSize =10
                    Name ="Label46"
                    Caption ="Unit"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =8520
                    LayoutCachedTop =600
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =900
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =6660
                    Top =120
                    Width =1050
                    FontSize =10
                    FontWeight =700
                    Name ="cmdEdit"
                    Caption ="Edit BOM"
                    OnClick ="mcrBOMEditOpen"
                    FontName ="Times New Roman"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =6660
                    LayoutCachedTop =120
                    LayoutCachedWidth =7710
                    LayoutCachedHeight =480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =7
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =3540
                    Top =120
                    Width =240
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =1
                    BackColor =255
                    Name ="PTCheckTotal"
                    ControlSource ="=Sum([PTCheck])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =120
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =420
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =2640
                    Top =780
                    Width =1380
                    Height =300
                    FontSize =10
                    Name ="Label52"
                    Caption ="Item"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =780
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =5460
                    Top =120
                    Width =1050
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="btnUpdateBOMUsage"
                    Caption ="Update Qty"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5460
                    LayoutCachedTop =120
                    LayoutCachedWidth =6510
                    LayoutCachedHeight =480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =336
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =60
                    Top =15
                    Width =480
                    Height =300
                    ColumnWidth =1185
                    ColumnOrder =10
                    FontSize =10
                    Name ="OrganizationCode"
                    ControlSource ="OrganizationCode"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =600
                    Top =15
                    Width =1740
                    Height =300
                    ColumnWidth =3000
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =1
                    Name ="AssemblyItemNumber"
                    ControlSource ="AssemblyItemNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2400
                    Top =15
                    Width =1620
                    Height =300
                    ColumnWidth =465
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =2
                    Name ="ComponentItemNumber"
                    ControlSource ="ComponentItemNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =4680
                    Top =15
                    Width =540
                    Height =300
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =3
                    Name ="OperationSeqNumber"
                    ControlSource ="OperationSeqNumber"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4680
                    LayoutCachedTop =15
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5280
                    Top =15
                    Width =900
                    Height =300
                    ColumnOrder =6
                    FontSize =10
                    TabIndex =4
                    Name ="EffectivityDate"
                    ControlSource ="EffectivityDate"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =15
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =7
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6240
                    Top =15
                    Width =900
                    Height =300
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =5
                    Name ="ComponentQuantity"
                    ControlSource ="ComponentQuantity"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =15
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7200
                    Top =15
                    Width =300
                    Height =300
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =6
                    Name ="CountPerUOM"
                    ControlSource ="CountPerUOM"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7200
                    LayoutCachedTop =15
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =4080
                    Top =15
                    Width =540
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =7
                    Name ="ItemSeqNumber"
                    ControlSource ="ItemSeqNumber"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =15
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =7
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =8640
                    Width =60
                    Height =300
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =8
                    BackColor =255
                    Name ="StoresPrefix"
                    ControlSource ="StoresPrefix"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8640
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =7
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =9060
                    Width =240
                    Height =300
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =9
                    BackColor =255
                    Name ="StoresSuffix"
                    ControlSource ="StoresSuffix"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =9060
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =8040
                    Top =15
                    Width =420
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =10
                    Name ="LayerID"
                    ControlSource ="LayerID"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =15
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    TextFontFamily =18
                    Left =8520
                    Top =15
                    Width =480
                    Height =300
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =11
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =15
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =315
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =7560
                    Top =15
                    Width =366
                    Height =306
                    TabIndex =12
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

                    LayoutCachedLeft =7560
                    LayoutCachedTop =15
                    LayoutCachedWidth =7926
                    LayoutCachedHeight =321
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =7
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    TextFontFamily =18
                    Left =7920
                    Width =120
                    Height =300
                    FontSize =10
                    TabIndex =13
                    BackColor =255
                    Name ="PTCheck"
                    ControlSource ="PTCheck"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7920
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            Height =0
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

Private Sub btnUpdateBOMUsage_Click()
    Call BomUsage
End Sub

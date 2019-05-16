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
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14235
    DatasheetFontHeight =10
    ItemSuffix =60
    Left =7005
    Top =1380
    Right =21495
    Bottom =12045
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc0a711c35633e540
    End
    RecordSource ="qryItemInterfaceBOMEditView"
    Caption ="Bill of Materials - EDIT"
    OnOpen ="[Event Procedure]"
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
            Height =1320
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Top =960
                    Width =480
                    Height =300
                    FontSize =10
                    Name ="Label15"
                    Caption ="Org"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =600
                    Top =960
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label16"
                    Caption ="Item"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2400
                    Top =960
                    Width =1440
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Value"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =4500
                    Top =720
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Label23"
                    Caption ="Oper"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =5100
                    Top =720
                    Width =900
                    Height =300
                    FontSize =10
                    Name ="Label25"
                    Caption ="Effectivity"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6060
                    Top =960
                    Width =900
                    Height =300
                    FontSize =10
                    Name ="Label27"
                    Caption ="Quantity"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7020
                    Top =960
                    Width =300
                    Height =300
                    FontSize =10
                    Name ="Label29"
                    Caption ="Ct"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =4500
                    Top =960
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Label31"
                    Caption ="Seq"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =5100
                    Top =960
                    Width =900
                    Height =300
                    FontSize =10
                    Name ="Label32"
                    Caption ="Date"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =60
                    Top =660
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label28"
                    Caption ="Bill of Materials"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =3900
                    Top =720
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Label34"
                    Caption ="Item"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =3900
                    Top =960
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Label35"
                    Caption ="Seq"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7800
                    Top =960
                    Width =480
                    Height =300
                    FontSize =10
                    Name ="Label44"
                    Caption ="Layer"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =7800
                    LayoutCachedTop =960
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =8340
                    Top =960
                    Width =360
                    Height =300
                    FontSize =10
                    Name ="Label46"
                    Caption ="Unit"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =8340
                    LayoutCachedTop =960
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1260
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =5940
                    Top =180
                    Width =1050
                    FontSize =10
                    FontWeight =700
                    Name ="cmdSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =7140
                    Top =180
                    Width =1050
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
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =4740
                    Top =180
                    Width =1050
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =4020
                    Top =180
                    Width =180
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =3
                    Name ="PTCheckTotal"
                    ControlSource ="=Sum([PTCheck])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =10620
                    Top =960
                    Width =1725
                    Height =300
                    FontSize =10
                    Name ="Label54"
                    Caption ="SupplySubinventory"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =10620
                    LayoutCachedTop =960
                    LayoutCachedWidth =12345
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =12420
                    Top =960
                    Width =1755
                    Height =300
                    FontSize =10
                    Name ="Label56"
                    Caption ="SupplyLocator"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =12420
                    LayoutCachedTop =960
                    LayoutCachedWidth =14175
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =8760
                    Top =960
                    Width =1785
                    Height =300
                    FontSize =10
                    Name ="Label58"
                    Caption ="SupplyType"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =8760
                    LayoutCachedTop =960
                    LayoutCachedWidth =10545
                    LayoutCachedHeight =1260
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
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Top =15
                    Width =480
                    Height =300
                    ColumnWidth =1185
                    ColumnOrder =10
                    FontSize =10
                    Name ="OrganizationCode"
                    ControlSource ="OrganizationCode"
                    DefaultValue ="=\"CAB\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Enabled = NotDefault
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
                    DefaultValue ="=[Forms]![frmMainMenu]![NewItemNumber]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2400
                    Top =15
                    Height =300
                    ColumnWidth =3180
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =2
                    Name ="ComponentItemNumber"
                    ControlSource ="ComponentItemNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =4500
                    Top =15
                    Width =540
                    Height =300
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =4
                    Name ="OperationSeqNumber"
                    ControlSource ="OperationSeqNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5100
                    Top =15
                    Width =900
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =5
                    Name ="EffectivityDate"
                    ControlSource ="EffectivityDate"
                    DefaultValue ="=Date()"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    DecimalPlaces =7
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6060
                    Top =15
                    Width =900
                    Height =300
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =6
                    Name ="ComponentQuantity"
                    ControlSource ="ComponentQuantity"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7020
                    Top =15
                    Width =300
                    Height =300
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =7
                    Name ="CountPerUOM"
                    ControlSource ="CountPerUOM"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =3900
                    Top =15
                    Width =540
                    Height =300
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =3
                    Name ="ItemSeqNumber"
                    ControlSource ="ItemSeqNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7800
                    Top =15
                    Width =480
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =8
                    Name ="LayerID"
                    ControlSource ="Layer"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =15
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =8340
                    Top =15
                    Width =360
                    Height =300
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =9
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =15
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =315
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7380
                    Top =15
                    Width =366
                    Height =306
                    TabIndex =10
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
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    TextFontFamily =18
                    Left =8700
                    Width =180
                    Height =300
                    FontSize =10
                    TabIndex =11
                    Name ="PTCheck"
                    ControlSource ="PTCheck"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8700
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =10620
                    Width =1740
                    Height =300
                    FontSize =10
                    TabIndex =12
                    Name ="Text55"
                    ControlSource ="SupplySubinventory"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =10620
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =12420
                    Width =1740
                    Height =300
                    FontSize =10
                    TabIndex =13
                    Name ="Text57"
                    ControlSource ="SupplyLocator"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =12420
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =1
                    TextFontFamily =18
                    Left =8760
                    Width =1800
                    Height =300
                    FontSize =10
                    TabIndex =14
                    Name ="Text59"
                    ControlSource ="SupplyType"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8760
                    LayoutCachedWidth =10560
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

Private Sub cmdRefresh_Click()
    DoCmd.RunMacro "mcrRequery"
    Call ReSequenceBOM(TempBOM)
End Sub

Private Sub cmdSave_Click()
    DoCmd.RunMacro "mcrBOMEditSave"
    Call BomUsage

End Sub

Private Sub Form_Open(Cancel As Integer)
    'If Me.ComponentItemNumber Like "pt*" Or Me.ComponentItemNumber Like "rbn*" Then
    'DoCmd.OpenForm "frmPremiseTBReplace"
    'End If

    'If CurrentProject.AllForms("frmPremiseTBReplace").IsLoaded Then
    'Forms("frmPremiseTBReplace").SetFocus
    'End If


End Sub

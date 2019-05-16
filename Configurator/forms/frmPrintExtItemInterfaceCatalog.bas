Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
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
    ItemSuffix =24
    Left =6675
    Top =1725
    Right =14355
    Bottom =8430
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x89a6285ed673e440
    End
    RecordSource ="qryPrintExtItemCatalogInterface"
    Caption ="Print Extension Catalog Values"
    DatasheetFontName ="Arial"
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
                    Left =4860
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
                    Left =6300
                    Top =60
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =2040
                    Top =15
                    Width =2760
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
                    Left =4860
                    Top =15
                    Width =2280
                    Height =300
                    ColumnWidth =465
                    FontSize =10
                    TabIndex =2
                    Name ="ElementValue"
                    ControlSource ="ElementValue"
                    OnEnter ="mcrOracleItemCatalogCheckOverride"
                    FontName ="Times New Roman"

                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =7200
                    Top =75
                    TabIndex =3
                    Name ="AllowOverride"
                    ControlSource ="AllowOverride"

                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =60
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End

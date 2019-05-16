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
    Width =5640
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =8565
    Top =3420
    Right =14205
    Bottom =6510
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x84fc0ac551ece440
    End
    RecordSource ="tblBOMBillOfMtlsInterface"
    Caption ="frmItemInterface"
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
                    TextAlign =2
                    TextFontFamily =18
                    Left =240
                    Top =420
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
                    Left =780
                    Top =420
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label16"
                    Caption ="Item"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =2520
                    Top =180
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Alt."
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =1860
                    Top =180
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label23"
                    Caption ="Assy"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    TextFontFamily =18
                    Left =1860
                    Top =420
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label31"
                    Caption ="Type"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =2520
                    Top =420
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label33"
                    Caption ="BOM"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =3180
                    Top =180
                    Width =840
                    Height =300
                    FontSize =10
                    Name ="Label35"
                    Caption ="Common"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =120
                    Top =120
                    Width =1740
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label29"
                    Caption ="BOM Interface"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =3180
                    Top =420
                    Width =840
                    Height =300
                    FontSize =10
                    Name ="Label38"
                    Caption ="Org"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =4080
                    Top =180
                    Width =1200
                    Height =300
                    FontSize =10
                    Name ="Label39"
                    Caption ="Common"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =4080
                    Top =420
                    Width =1200
                    Height =300
                    FontSize =10
                    Name ="Label40"
                    Caption ="Item"
                    FontName ="Times New Roman"
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
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =240
                    Top =15
                    Width =480
                    Height =300
                    ColumnWidth =1185
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
                    Left =780
                    Top =15
                    Width =1020
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =1
                    Name ="AssemblyItemNumber"
                    ControlSource ="ItemNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =2520
                    Top =15
                    Width =600
                    Height =300
                    ColumnWidth =465
                    FontSize =10
                    TabIndex =2
                    Name ="AlternateBOMDesignator"
                    ControlSource ="AlternateBOMDesignator"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =1860
                    Top =15
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="AssemblyType"
                    ControlSource ="AssemblyType"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =3180
                    Width =840
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="CommonOrgCode"
                    ControlSource ="CommonOrgCode"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =4080
                    Width =1200
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="CommonItemNumber"
                    ControlSource ="CommonItemNumber"
                    Format ="Standard"
                    FontName ="Times New Roman"

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

Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6480
    DatasheetFontHeight =10
    ItemSuffix =80
    Left =15810
    Right =19050
    Bottom =2835
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x03209f1e4bdee440
    End
    RecordSource ="qryPremiseItemDetails"
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
            Height =60
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3000
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1500
                    Top =60
                    Width =2040
                    Height =300
                    ColumnWidth =1185
                    ColumnOrder =10
                    FontSize =10
                    Name ="OracleItemNumber"
                    ControlSource ="OracleItemNumber"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =60
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            TextFontFamily =18
                            Left =3720
                            Top =1860
                            Width =810
                            Height =300
                            FontSize =10
                            Name ="Label72"
                            Caption ="Base"
                            FontName ="Times New Roman"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =1860
                            LayoutCachedWidth =4530
                            LayoutCachedHeight =2160
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =10
                    Name ="Label15"
                    Caption ="Base Item"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1500
                    Top =420
                    Width =2040
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="DesignCode"
                    ControlSource ="DesignCode"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =420
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =420
                    Width =1320
                    Height =300
                    FontSize =10
                    Name ="Label50"
                    Caption ="Design Code"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1500
                    Top =780
                    Width =2040
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Customer"
                    ControlSource ="Customer"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =780
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =780
                    Width =1320
                    Height =300
                    FontSize =10
                    Name ="Label52"
                    Caption ="Customer"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1500
                    Top =1140
                    Width =2040
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="CustomerPartNumber"
                    ControlSource ="CustomerPartNumber"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1440
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =1140
                    Width =1320
                    Height =300
                    FontSize =10
                    Name ="Label54"
                    Caption ="Customer Part"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1500
                    Top =1500
                    Width =2040
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="FiberType1"
                    ControlSource ="FiberType1"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1500
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1800
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =1500
                    Width =1320
                    Height =300
                    FontSize =10
                    Name ="Label56"
                    Caption ="Fiber Type 1"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1500
                    Top =1860
                    Width =2040
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="FiberType2"
                    ControlSource ="FiberType2"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1860
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =2160
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =1860
                    Width =1320
                    Height =300
                    FontSize =10
                    Name ="Label58"
                    Caption ="Fiber Type 2"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1500
                    Top =2220
                    Width =2040
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="FiberType3"
                    ControlSource ="FiberType3"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =2220
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =2220
                    Width =1320
                    Height =300
                    FontSize =10
                    Name ="Label60"
                    Caption ="Fiber Type 3"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4560
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="JacketMaterial"
                    ControlSource ="JacketMaterial"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =60
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =360
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =3840
                    Top =60
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label62"
                    Caption ="Jacket"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4560
                    Top =420
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="JacketColor"
                    ControlSource ="JacketColor"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =420
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =3840
                    Top =420
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label64"
                    Caption ="Color"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3840
                    LayoutCachedTop =420
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4560
                    Top =780
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =9
                    Name ="ColorChip"
                    ControlSource ="ColorChip"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =780
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =3840
                    Top =780
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label66"
                    Caption ="Chip"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3840
                    LayoutCachedTop =780
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =1080
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4560
                    Top =1140
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =10
                    Name ="Routing1"
                    ControlSource ="Routing1"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =1140
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1440
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =3720
                    Top =1140
                    Width =810
                    Height =300
                    FontSize =10
                    Name ="Label68"
                    Caption ="Routing1"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =1440
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4560
                    Top =1500
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =11
                    Name ="Routing2"
                    ControlSource ="Routing2"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =1500
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1800
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =3720
                    Top =1500
                    Width =810
                    Height =300
                    FontSize =10
                    Name ="Label70"
                    Caption ="Routing2"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =1500
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =1800
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4560
                    Top =1860
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =12
                    Name ="Base"
                    ControlSource ="Base"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =2160
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4560
                    Top =2220
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =13
                    Name ="Location"
                    ControlSource ="CablePasses"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =2220
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =2520
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            TextFontFamily =18
                            Left =3720
                            Top =2220
                            Width =810
                            Height =300
                            FontSize =10
                            Name ="Label73"
                            Caption ="Passes"
                            FontName ="Times New Roman"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =2220
                            LayoutCachedWidth =4530
                            LayoutCachedHeight =2520
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1500
                    Top =2580
                    Width =2040
                    Height =300
                    FontSize =10
                    TabIndex =14
                    Name ="UnitSeries"
                    ControlSource ="qryPremiseItemDetails_Union.[Unit Series]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =2580
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =2880
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =2580
                    Width =1320
                    Height =300
                    FontSize =10
                    Name ="Label76"
                    Caption ="Unit Series"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4560
                    Top =2580
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =15
                    Name ="Listing"
                    ControlSource ="qryPremiseItemDetails_Union.Listing"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =2580
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =2880
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            TextFontFamily =18
                            Left =3720
                            Top =2580
                            Width =810
                            Height =300
                            FontSize =10
                            Name ="Label78"
                            Caption ="Listing"
                            FontName ="Times New Roman"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =2580
                            LayoutCachedWidth =4530
                            LayoutCachedHeight =2880
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =3
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =5880
                    Top =780
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =16
                    Name ="Text79"
                    ControlSource ="Color_Chip_Concentration"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =780
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =1080
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

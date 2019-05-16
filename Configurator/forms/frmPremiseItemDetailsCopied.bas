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
    Width =5640
    DatasheetFontHeight =10
    ItemSuffix =68
    Left =15810
    Top =4800
    Right =19050
    Bottom =7515
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa4e09f1e4bdee440
    End
    RecordSource ="qryPremiseItemCopied"
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label67"
                    Caption ="COPIED ITEM"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            Height =2700
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
                    Left =1440
                    Top =60
                    Width =1860
                    Height =300
                    ColumnWidth =1185
                    ColumnOrder =10
                    FontSize =10
                    Name ="OracleItemNumber"
                    ControlSource ="OracleItemNumber"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =10
                    Name ="Label15"
                    Caption ="Copied Item"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1440
                    Top =420
                    Width =1860
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="DesignCode"
                    ControlSource ="DesignCode"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
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
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1440
                    Top =780
                    Width =1860
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Customer"
                    ControlSource ="Customer"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
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
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1440
                    Top =1140
                    Width =1860
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="CustomerPartNumber"
                    ControlSource ="CustomerPartNumber"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
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
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1440
                    Top =1500
                    Width =1860
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="FiberType1"
                    ControlSource ="FiberType1"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
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
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1440
                    Top =1860
                    Width =1860
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="FiberType2"
                    ControlSource ="FiberType2"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
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
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =1440
                    Top =2220
                    Width =1860
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="FiberType3"
                    ControlSource ="FiberType3"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
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
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4320
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="JacketMaterial"
                    ControlSource ="JacketMaterial"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    TextFontFamily =18
                    Left =3660
                    Top =60
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label62"
                    Caption ="Jacket"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4320
                    Top =420
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="JacketColor"
                    ControlSource ="JacketColor"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    TextFontFamily =18
                    Left =3660
                    Top =420
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label64"
                    Caption ="Color"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4320
                    Top =780
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =9
                    Name ="ColorChip"
                    ControlSource ="ColorChip"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    TextFontFamily =18
                    Left =3660
                    Top =780
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label66"
                    Caption ="Chip"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =4320
                    Top =1080
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =10
                    Name ="Text67"
                    ControlSource ="Routing1"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    TextFontFamily =18
                    Left =3480
                    Top =1080
                    Width =810
                    Height =300
                    FontSize =10
                    Name ="Label68"
                    Caption ="Routing1"
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
                    Left =4320
                    Top =1440
                    Width =1320
                    Height =300
                    FontSize =10
                    TabIndex =11
                    Name ="Text69"
                    ControlSource ="Routing2"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =3480
                    Top =1440
                    Width =810
                    Height =300
                    FontSize =10
                    Name ="Label70"
                    Caption ="Routing2"
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

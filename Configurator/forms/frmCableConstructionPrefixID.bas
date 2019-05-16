Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8820
    DatasheetFontHeight =11
    ItemSuffix =53
    Left =345
    Top =4770
    Right =9195
    Bottom =11130
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xe07981f2aaffe440
    End
    RecordSource ="CableRunSpeeds"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x000000000000000000000000000000000000000062250000101d000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    Moveable =0
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =-2147483643
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            ForeThemeColorIndex =2
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            KeepTogether = NotDefault
            Height =5760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =223
                    TextAlign =1
                    Left =4471
                    Top =780
                    Width =1766
                    Height =260
                    FontSize =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NominalWall"
                    ControlSource ="Process Status"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4471
                    LayoutCachedTop =780
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =1040
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =4471
                            Top =540
                            Width =1766
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Nominal Wall Label"
                            Caption ="Process Status"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Nominal_Wall_Label"
                            LayoutCachedLeft =4471
                            LayoutCachedTop =540
                            LayoutCachedWidth =6237
                            LayoutCachedHeight =800
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =300
                    Top =2060
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MaxAveWall"
                    ControlSource ="Set-up time level 1"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =300
                    LayoutCachedTop =2060
                    LayoutCachedWidth =1781
                    LayoutCachedHeight =2320
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =304
                            Top =1802
                            Width =1500
                            Height =255
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Max Ave Wall Label"
                            Caption ="Set-up Time Level 1"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Max_Ave_Wall_Label"
                            LayoutCachedLeft =304
                            LayoutCachedTop =1802
                            LayoutCachedWidth =1804
                            LayoutCachedHeight =2057
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =1781
                    Top =2060
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MinAveWall"
                    ControlSource ="Set-up time level 2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1781
                    LayoutCachedTop =2060
                    LayoutCachedWidth =3262
                    LayoutCachedHeight =2320
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =1
                            Left =1781
                            Top =1800
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Min Ave Wall Label"
                            Caption ="Set-up time Level 2"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Min_Ave_Wall_Label"
                            LayoutCachedLeft =1781
                            LayoutCachedTop =1800
                            LayoutCachedWidth =3262
                            LayoutCachedHeight =2060
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =4638
                    Top =2660
                    Width =1500
                    Height =260
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MinSpotWall"
                    ControlSource ="Buff#11 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4638
                    LayoutCachedTop =2660
                    LayoutCachedWidth =6138
                    LayoutCachedHeight =2920
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =4635
                            Top =2400
                            Width =1500
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Min Spot Wall Label"
                            Caption ="Buff#11 Line Speed"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Min_Spot_Wall_Label"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =2400
                            LayoutCachedWidth =6135
                            LayoutCachedHeight =2660
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Left =120
                    Top =3060
                    Width =8504
                    Name ="Line130"
                    LayoutCachedLeft =120
                    LayoutCachedTop =3060
                    LayoutCachedWidth =8624
                    LayoutCachedHeight =3060
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =223
                    TextAlign =1
                    Left =1680
                    Top =780
                    Width =1321
                    Height =260
                    FontSize =8
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Jacket Material"
                    ControlSource ="Buffering Factor"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Jacket_Material"
                    AsianLineBreak =0

                    LayoutCachedLeft =1680
                    LayoutCachedTop =780
                    LayoutCachedWidth =3001
                    LayoutCachedHeight =1040
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1680
                            Top =540
                            Width =1321
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Jacket Material Label"
                            Caption ="Buffering Factor"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Jacket_Material_Label"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =540
                            LayoutCachedWidth =3001
                            LayoutCachedHeight =800
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =240
                    Top =780
                    Height =260
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txPrefixID"
                    ControlSource ="PrefixID"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =240
                    LayoutCachedTop =780
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =1040
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =240
                            Top =540
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label845"
                            Caption ="PrefixID:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =240
                            LayoutCachedTop =540
                            LayoutCachedWidth =1680
                            LayoutCachedHeight =800
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Left =120
                    Top =1140
                    Width =8504
                    Name ="Line0"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1140
                    LayoutCachedWidth =8624
                    LayoutCachedHeight =1140
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =215
                    TextAlign =1
                    Left =3001
                    Top =780
                    Width =1470
                    Height =255
                    FontSize =8
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Chip Type"
                    ControlSource ="Date Adjustments"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Chip_Type"
                    AsianLineBreak =0

                    LayoutCachedLeft =3001
                    LayoutCachedTop =780
                    LayoutCachedWidth =4471
                    LayoutCachedHeight =1035
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =3001
                            Top =540
                            Width =1470
                            Height =255
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Chip Type Label"
                            Caption ="Date Adjustments"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Chip_Type_Label"
                            LayoutCachedLeft =3001
                            LayoutCachedTop =540
                            LayoutCachedWidth =4471
                            LayoutCachedHeight =795
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =120
                    Width =3000
                    Height =405
                    FontSize =14
                    FontWeight =600
                    Name ="Label55"
                    Caption ="Run Speed Data"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =405
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =3262
                    Top =2060
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text1"
                    ControlSource ="Buff#1 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3262
                    LayoutCachedTop =2060
                    LayoutCachedWidth =4743
                    LayoutCachedHeight =2320
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =3262
                            Top =1800
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label2"
                            Caption ="Buff#1 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3262
                            LayoutCachedTop =1800
                            LayoutCachedWidth =4743
                            LayoutCachedHeight =2060
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =6224
                    Top =2060
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text3"
                    ControlSource ="Buff#4 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6224
                    LayoutCachedTop =2060
                    LayoutCachedWidth =7705
                    LayoutCachedHeight =2320
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =6224
                            Top =1800
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label4"
                            Caption ="Buff#4 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6224
                            LayoutCachedTop =1800
                            LayoutCachedWidth =7705
                            LayoutCachedHeight =2060
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =1676
                    Top =2660
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text5"
                    ControlSource ="Buff#6 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1676
                    LayoutCachedTop =2660
                    LayoutCachedWidth =3157
                    LayoutCachedHeight =2920
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1676
                            Top =2400
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label6"
                            Caption ="Buff#6 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1676
                            LayoutCachedTop =2400
                            LayoutCachedWidth =3157
                            LayoutCachedHeight =2660
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =6138
                    Top =2660
                    Width =1500
                    Height =260
                    FontSize =8
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text7"
                    ControlSource ="B18_Line_Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6138
                    LayoutCachedTop =2660
                    LayoutCachedWidth =7638
                    LayoutCachedHeight =2920
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =6131
                            Top =2400
                            Width =1500
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label8"
                            Caption ="Buff#18 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6131
                            LayoutCachedTop =2400
                            LayoutCachedWidth =7631
                            LayoutCachedHeight =2660
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =4743
                    Top =2060
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text9"
                    ControlSource ="Buff#2 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4743
                    LayoutCachedTop =2060
                    LayoutCachedWidth =6224
                    LayoutCachedHeight =2320
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =4743
                            Top =1800
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label10"
                            Caption ="Buff#2 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =4743
                            LayoutCachedTop =1800
                            LayoutCachedWidth =6224
                            LayoutCachedHeight =2060
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =3157
                    Top =2660
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text11"
                    ControlSource ="Buff#9 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3157
                    LayoutCachedTop =2660
                    LayoutCachedWidth =4638
                    LayoutCachedHeight =2920
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =3157
                            Top =2400
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label12"
                            Caption ="Buff#9 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3157
                            LayoutCachedTop =2400
                            LayoutCachedWidth =4638
                            LayoutCachedHeight =2660
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =247
                    Top =4043
                    Width =1500
                    Height =260
                    FontSize =8
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text13"
                    ControlSource ="Buff#10 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =247
                    LayoutCachedTop =4043
                    LayoutCachedWidth =1747
                    LayoutCachedHeight =4303
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =240
                            Top =3780
                            Width =1500
                            Height =270
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label14"
                            Caption ="Buff#10 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =240
                            LayoutCachedTop =3780
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =4050
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =243
                    Top =4643
                    Width =1856
                    Height =260
                    FontSize =8
                    TabIndex =14
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text15"
                    ControlSource ="Buff#11 Line Speed (SH)"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =243
                    LayoutCachedTop =4643
                    LayoutCachedWidth =2099
                    LayoutCachedHeight =4903
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =240
                            Top =4380
                            Width =1860
                            Height =270
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label16"
                            Caption ="Buff#11 Line Speed (SH)"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4380
                            LayoutCachedWidth =2100
                            LayoutCachedHeight =4650
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =2115
                    Top =4640
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =15
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text17"
                    ControlSource ="SH#3 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2115
                    LayoutCachedTop =4640
                    LayoutCachedWidth =3596
                    LayoutCachedHeight =4900
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2115
                            Top =4380
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label18"
                            Caption ="SH#3 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =2115
                            LayoutCachedTop =4380
                            LayoutCachedWidth =3596
                            LayoutCachedHeight =4640
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =3596
                    Top =4640
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =16
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text19"
                    ControlSource ="SH#4 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3596
                    LayoutCachedTop =4640
                    LayoutCachedWidth =5077
                    LayoutCachedHeight =4900
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =3596
                            Top =4380
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label20"
                            Caption ="SH#4 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3596
                            LayoutCachedTop =4380
                            LayoutCachedWidth =5077
                            LayoutCachedHeight =4640
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =240
                    Top =5300
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =17
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text21"
                    ControlSource ="C#4 Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =240
                    LayoutCachedTop =5300
                    LayoutCachedWidth =1721
                    LayoutCachedHeight =5560
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =240
                            Top =5040
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label22"
                            Caption ="C#8 Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =240
                            LayoutCachedTop =5040
                            LayoutCachedWidth =1721
                            LayoutCachedHeight =5300
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =3221
                    Top =5300
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =18
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text27"
                    ControlSource ="PlannedLotSize"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3221
                    LayoutCachedTop =5300
                    LayoutCachedWidth =4702
                    LayoutCachedHeight =5560
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =3221
                            Top =5040
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label28"
                            Caption ="PlannedLotSize"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3221
                            LayoutCachedTop =5040
                            LayoutCachedWidth =4702
                            LayoutCachedHeight =5300
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =4702
                    Top =5300
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =19
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text29"
                    ControlSource ="shrinkage limit"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4702
                    LayoutCachedTop =5300
                    LayoutCachedWidth =6183
                    LayoutCachedHeight =5560
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =4702
                            Top =5040
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label30"
                            Caption ="shrinkage limit"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =4702
                            LayoutCachedTop =5040
                            LayoutCachedWidth =6183
                            LayoutCachedHeight =5300
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =1620
                    Top =3380
                    Width =1695
                    Height =260
                    FontSize =8
                    TabIndex =20
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text31"
                    ControlSource ="Sheathing Line"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3380
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =3640
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1620
                            Top =3120
                            Width =1695
                            Height =255
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label32"
                            Caption ="Primary Sheathing Line"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1620
                            LayoutCachedTop =3120
                            LayoutCachedWidth =3315
                            LayoutCachedHeight =3375
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =1740
                    Top =5300
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =22
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text23"
                    ControlSource ="Armor Line Speed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1740
                    LayoutCachedTop =5300
                    LayoutCachedWidth =3221
                    LayoutCachedHeight =5560
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =1740
                            Top =5040
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label24"
                            Caption ="Armor Line Speed"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1740
                            LayoutCachedTop =5040
                            LayoutCachedWidth =3221
                            LayoutCachedHeight =5300
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =1743
                    Top =4043
                    Width =1856
                    Height =260
                    FontSize =8
                    TabIndex =24
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text45"
                    ControlSource ="Buff#9 Line Speed (SH)"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1743
                    LayoutCachedTop =4043
                    LayoutCachedWidth =3599
                    LayoutCachedHeight =4303
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =1740
                            Top =3780
                            Width =1860
                            Height =270
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label46"
                            Caption ="Buff#9 Line Speed (SH)"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1740
                            LayoutCachedTop =3780
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =4050
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =215
                    TextAlign =1
                    Left =1680
                    Top =1440
                    Width =1256
                    Height =260
                    FontSize =8
                    TabIndex =23
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text41"
                    ControlSource ="PrimaryBufferingLine"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1680
                    LayoutCachedTop =1440
                    LayoutCachedWidth =2936
                    LayoutCachedHeight =1700
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1680
                            Top =1200
                            Width =1256
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label42"
                            Caption ="Primary Buff"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2936
                            LayoutCachedHeight =1460
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =300
                    Top =2660
                    Width =1376
                    Height =260
                    FontSize =8
                    TabIndex =21
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text33"
                    ControlSource ="Interlock"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =300
                    LayoutCachedTop =2660
                    LayoutCachedWidth =1676
                    LayoutCachedHeight =2920
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =300
                            Top =2400
                            Width =1376
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label34"
                            Caption ="Interlock"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =300
                            LayoutCachedTop =2400
                            LayoutCachedWidth =1676
                            LayoutCachedHeight =2660
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6660
                    Width =720
                    FontSize =8
                    TabIndex =25
                    ForeColor =-2147483630
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =6660
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =16777215
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =3600
                    Top =4040
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =26
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text47"
                    ControlSource ="B2_Second_Pass"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3600
                    LayoutCachedTop =4040
                    LayoutCachedWidth =5081
                    LayoutCachedHeight =4300
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =3600
                            Top =3780
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label48"
                            Caption ="B2 2nd Pass"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3600
                            LayoutCachedTop =3780
                            LayoutCachedWidth =5081
                            LayoutCachedHeight =4040
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =1200
                    Width =1320
                    Height =405
                    FontSize =14
                    FontWeight =600
                    Name ="Label49"
                    Caption ="1st Pass"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =1605
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =3120
                    Width =1380
                    Height =405
                    FontSize =14
                    FontWeight =600
                    Name ="Label50"
                    Caption ="2nd Pass"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =120
                    LayoutCachedTop =3120
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =3525
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =5100
                    Top =4040
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =27
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text51"
                    ControlSource ="B18_Line_Speed_Second_Pass"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5100
                    LayoutCachedTop =4040
                    LayoutCachedWidth =6581
                    LayoutCachedHeight =4300
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =5100
                            Top =3780
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label52"
                            Caption ="B18 2nd Pass"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5100
                            LayoutCachedTop =3780
                            LayoutCachedWidth =6581
                            LayoutCachedHeight =4040
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
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

Private Sub btnMils_Click()
On Error GoTo Error_Handler

DoCmd.OpenForm "frmMilsConversion"

Forms!frmMilsConversion!NominalWall = Round(CDbl(Me.NominalWall) * 39.37, 1)
Forms!frmMilsConversion!MinAveWall = Round(CDbl(Me.MinAveWall) * 39.37, 1)
Forms!frmMilsConversion!MinSpotWall = Round(CDbl(Me.MinSpotWall) * 39.37, 1)
Forms!frmMilsConversion!MaxAveWall = Round(CDbl(Me.MaxAveWall) * 39.37, 1)

Error_Exit:
Exit Sub

Error_Handler:
'MsgBox Err.Description & " " & Err.Number
Resume Error_Exit

End Sub

Private Sub btnEdit_Click()
Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If ctl.Locked <> False Then
            ctl.Locked = False
        End If
    End If
Next ctl
Me.txPrefixID.Locked = True
End Sub

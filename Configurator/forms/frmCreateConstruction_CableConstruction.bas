Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
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
    Width =9570
    DatasheetFontHeight =11
    ItemSuffix =53
    Left =12780
    Top =1080
    Right =22350
    Bottom =11295
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x179796e15503e540
    End
    RecordSource ="tblCableConstructions_Local"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x0000000000000000000000000000000000000000622500006824000001000000 ,
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
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            KeepTogether = NotDefault
            Height =10100
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =1721
                    Top =2700
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Nominal OD"
                    ControlSource ="NominalOD"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Nominal_OD"
                    AsianLineBreak =0

                    LayoutCachedLeft =1721
                    LayoutCachedTop =2700
                    LayoutCachedWidth =3202
                    LayoutCachedHeight =2960
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
                            Left =540
                            Top =2700
                            Width =1016
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Nominal OD Label"
                            Caption ="Nominal OD"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Nominal_OD_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =2700
                            LayoutCachedWidth =1556
                            LayoutCachedHeight =2960
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
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =4800
                    Top =2700
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="OD Tol  (+)"
                    ControlSource ="ODTolPlus"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="OD_Tol_____"
                    AsianLineBreak =0

                    LayoutCachedLeft =4800
                    LayoutCachedTop =2700
                    LayoutCachedWidth =6281
                    LayoutCachedHeight =2960
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =7680
                    Top =2700
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =14
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="OD Tol  (-)1"
                    ControlSource ="ODTolMinus"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="OD_Tol_____1"
                    AsianLineBreak =0

                    LayoutCachedLeft =7680
                    LayoutCachedTop =2700
                    LayoutCachedWidth =9161
                    LayoutCachedHeight =2960
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =1721
                    Top =2960
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =15
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Height"
                    ControlSource ="Height"
                    Format ="General Number"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1721
                    LayoutCachedTop =2960
                    LayoutCachedWidth =3202
                    LayoutCachedHeight =3220
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
                            OverlapFlags =87
                            TextAlign =1
                            Left =540
                            Top =2960
                            Width =1016
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Height Label"
                            Caption ="Height"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Height_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =2960
                            LayoutCachedWidth =1556
                            LayoutCachedHeight =3220
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
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =4800
                    Top =2960
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =16
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Height Tol (+)"
                    ControlSource ="HeightTolPlus"
                    Format ="General Number"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Height_Tol____"
                    AsianLineBreak =0

                    LayoutCachedLeft =4800
                    LayoutCachedTop =2960
                    LayoutCachedWidth =6281
                    LayoutCachedHeight =3220
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =7680
                    Top =2960
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =17
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Height Tol (-)2"
                    ControlSource ="HeightTolMinus"
                    Format ="General Number"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Height_Tol____2"
                    AsianLineBreak =0

                    LayoutCachedLeft =7680
                    LayoutCachedTop =2960
                    LayoutCachedWidth =9161
                    LayoutCachedHeight =3220
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =540
                    Top =3558
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =19
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txNominalWall"
                    ControlSource ="NominalWall"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =540
                    LayoutCachedTop =3558
                    LayoutCachedWidth =2021
                    LayoutCachedHeight =3818
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
                            Left =540
                            Top =3298
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Nominal Wall Label"
                            Caption ="Nominal Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Nominal_Wall_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =3298
                            LayoutCachedWidth =2021
                            LayoutCachedHeight =3558
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
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =2021
                    Top =3558
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =20
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txMaxAveWall"
                    ControlSource ="MaxAveWall"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2021
                    LayoutCachedTop =3558
                    LayoutCachedWidth =3502
                    LayoutCachedHeight =3818
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
                            Left =2021
                            Top =3298
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Max Ave Wall Label"
                            Caption ="Max Ave Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Max_Ave_Wall_Label"
                            LayoutCachedLeft =2021
                            LayoutCachedTop =3298
                            LayoutCachedWidth =3502
                            LayoutCachedHeight =3558
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
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =3502
                    Top =3558
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =21
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txMinAveWall"
                    ControlSource ="MinAveWall"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3502
                    LayoutCachedTop =3558
                    LayoutCachedWidth =4983
                    LayoutCachedHeight =3818
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
                            Left =3502
                            Top =3298
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Min Ave Wall Label"
                            Caption ="Min Ave Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Min_Ave_Wall_Label"
                            LayoutCachedLeft =3502
                            LayoutCachedTop =3298
                            LayoutCachedWidth =4983
                            LayoutCachedHeight =3558
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
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =4983
                    Top =3558
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =22
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txMinSpotWall"
                    ControlSource ="MinSpotWall"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4983
                    LayoutCachedTop =3558
                    LayoutCachedWidth =6464
                    LayoutCachedHeight =3818
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
                            Left =4983
                            Top =3298
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Min Spot Wall Label"
                            Caption ="Min Spot Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Min_Spot_Wall_Label"
                            LayoutCachedLeft =4983
                            LayoutCachedTop =3298
                            LayoutCachedWidth =6464
                            LayoutCachedHeight =3558
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
                    Left =480
                    Top =3960
                    Width =8504
                    Name ="Line130"
                    LayoutCachedLeft =480
                    LayoutCachedTop =3960
                    LayoutCachedWidth =8984
                    LayoutCachedHeight =3960
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Left =480
                    Top =2580
                    Width =8504
                    Name ="Line0"
                    LayoutCachedLeft =480
                    LayoutCachedTop =2580
                    LayoutCachedWidth =8984
                    LayoutCachedHeight =2580
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Left =420
                    Top =8280
                    Width =8504
                    Name ="Line135"
                    LayoutCachedLeft =420
                    LayoutCachedTop =8280
                    LayoutCachedWidth =8924
                    LayoutCachedHeight =8280
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =1901
                    Top =8400
                    Width =6617
                    Height =860
                    FontSize =8
                    TabIndex =53
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text0"
                    ControlSource ="Instructions"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1901
                    LayoutCachedTop =8400
                    LayoutCachedWidth =8518
                    LayoutCachedHeight =9260
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
                            OverlapFlags =85
                            TextAlign =1
                            Left =420
                            Top =8400
                            Width =1320
                            Height =540
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label1"
                            Caption ="Construction Instructions"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =420
                            LayoutCachedTop =8400
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =8940
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    TextAlign =1
                    Left =1901
                    Top =9240
                    Width =6617
                    Height =860
                    FontSize =8
                    TabIndex =55
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text14"
                    ControlSource ="Instructions2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1901
                    LayoutCachedTop =9240
                    LayoutCachedWidth =8518
                    LayoutCachedHeight =10100
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
                            OverlapFlags =85
                            TextAlign =1
                            Left =420
                            Top =9240
                            Width =1320
                            Height =540
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label15"
                            Caption ="Construction Instructions"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =420
                            LayoutCachedTop =9240
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =9780
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =1440
                    Top =4080
                    Width =2310
                    Height =260
                    FontSize =8
                    TabIndex =23
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="FRP Dia"
                    ControlSource ="FRP_Dia"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="FRP_Dia"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4080
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =4340
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
                            Left =540
                            Top =4080
                            Width =780
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="FRP Dia Label"
                            Caption ="FRP Dia"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="FRP_Dia_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =4080
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =4340
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    TextAlign =1
                    Left =1440
                    Top =4320
                    Width =2310
                    Height =260
                    FontSize =8
                    TabIndex =27
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM Type"
                    ControlSource ="CM"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_Type"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4320
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =4580
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
                            Left =540
                            Top =4335
                            Width =780
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="CM Type Label"
                            Caption ="CM Type"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="CM_Type_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =4335
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =4595
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    Left =1440
                    Top =4620
                    Width =2310
                    Height =260
                    FontSize =8
                    TabIndex =30
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM Matrl"
                    ControlSource ="CMMaterial"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_Matrl"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4620
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =4880
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
                            OverlapFlags =87
                            Left =540
                            Top =4590
                            Width =780
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="CM Matrl Label"
                            Caption ="CM Matrl"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="CM_Matrl_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =4590
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =4850
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =223
                    TextAlign =2
                    Left =5040
                    Top =4080
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =24
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM OD"
                    ControlSource ="CM_OD"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_OD"
                    AsianLineBreak =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =4080
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =4340
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
                            Left =3960
                            Top =4080
                            Width =1155
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="CM OD Label"
                            Caption ="CM OD"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="CM_OD_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =4080
                            LayoutCachedWidth =5115
                            LayoutCachedHeight =4340
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =2
                    Left =6240
                    Top =4080
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =25
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM OD Tol +"
                    ControlSource ="CMODTolPlus"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_OD_Tol__"
                    AsianLineBreak =0

                    LayoutCachedLeft =6240
                    LayoutCachedTop =4080
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =4340
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    TextAlign =2
                    Left =7380
                    Top =4080
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =26
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM OD Tol -1"
                    ControlSource ="CMODTolMinus"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_OD_Tol__1"
                    AsianLineBreak =0

                    LayoutCachedLeft =7380
                    LayoutCachedTop =4080
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =4340
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    TextAlign =2
                    Left =5040
                    Top =4320
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =28
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM Wall"
                    ControlSource ="CMWall"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_Wall"
                    AsianLineBreak =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =4320
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =4580
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
                            Left =3960
                            Top =4335
                            Width =1155
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="CM Wall Label"
                            Caption ="CM Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="CM_Wall_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =4335
                            LayoutCachedWidth =5115
                            LayoutCachedHeight =4595
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =2
                    Left =5040
                    Top =4620
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =31
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM Min Wall"
                    ControlSource ="CMMinWall"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_Min_Wall"
                    AsianLineBreak =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =4620
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =4880
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
                            Left =3960
                            Top =4590
                            Width =1080
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="CM Min Wall Label"
                            Caption ="CM Min Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="CM_Min_Wall_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =4590
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =4850
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    Left =1440
                    Top =4980
                    Width =2295
                    Height =260
                    FontSize =8
                    TabIndex =33
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Binder #1"
                    ControlSource ="Binder1"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Binder__1"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4980
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =5240
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
                            Left =540
                            Top =4980
                            Width =780
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Binder #1 Label"
                            Caption ="Binder #1"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Binder__1_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =4980
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =5240
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =2
                    Left =5010
                    Top =4980
                    Width =1050
                    Height =260
                    FontSize =8
                    TabIndex =34
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Qty Binder #1"
                    ControlSource ="QtyBinder1"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Qty_Binder__1"
                    AsianLineBreak =0

                    LayoutCachedLeft =5010
                    LayoutCachedTop =4980
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =5240
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
                            Left =3960
                            Top =4980
                            Width =1050
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Qty Binder #1 Label"
                            Caption ="Qty Binder #1"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Qty_Binder__1_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =4980
                            LayoutCachedWidth =5010
                            LayoutCachedHeight =5240
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =1440
                    Top =5220
                    Width =2295
                    Height =260
                    FontSize =8
                    TabIndex =36
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Binder #2"
                    ControlSource ="Binder2"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Binder__2"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =5220
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =5480
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
                            OverlapFlags =87
                            Left =540
                            Top =5235
                            Width =780
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Binder #2 Label"
                            Caption ="Binder #2"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Binder__2_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =5235
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =5495
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =2
                    Left =5010
                    Top =5235
                    Width =1050
                    Height =260
                    FontSize =8
                    TabIndex =37
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Qty Binder #2"
                    ControlSource ="QtyBinder2"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Qty_Binder__2"
                    AsianLineBreak =0

                    LayoutCachedLeft =5010
                    LayoutCachedTop =5235
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =5495
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
                            Left =3960
                            Top =5235
                            Width =1050
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Qty Binder #2 Label"
                            Caption ="Qty Binder #2"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Qty_Binder__2_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =5235
                            LayoutCachedWidth =5010
                            LayoutCachedHeight =5495
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1440
                    Top =5580
                    Width =2295
                    Height =260
                    FontSize =8
                    TabIndex =39
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Core Wrap"
                    ControlSource ="CoreWrap"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Core_Wrap"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =5580
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =5840
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
                            OverlapFlags =85
                            Left =540
                            Top =5580
                            Width =780
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Core Wrap Label"
                            Caption ="Core Wrap"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Core_Wrap_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =5580
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =5840
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1440
                    Top =5880
                    Width =2295
                    Height =260
                    FontSize =8
                    TabIndex =41
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Binder #3"
                    ControlSource ="Binder3"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Binder__3"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =5880
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =6140
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
                            OverlapFlags =85
                            Left =540
                            Top =5880
                            Width =780
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Binder #3 Label"
                            Caption ="Binder #3"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Binder__3_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =5880
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =6140
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =2
                    Left =5010
                    Top =5880
                    Width =1050
                    Height =260
                    FontSize =8
                    TabIndex =42
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Qty Binder #3"
                    ControlSource ="QtyBinder3"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Qty_Binder__3"
                    AsianLineBreak =0

                    LayoutCachedLeft =5010
                    LayoutCachedTop =5880
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =6140
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
                            Left =3960
                            Top =5880
                            Width =1050
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Qty Binder #3 Label"
                            Caption ="Qty Binder #3"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Qty_Binder__3_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =5880
                            LayoutCachedWidth =5010
                            LayoutCachedHeight =6140
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
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1440
                    Top =6240
                    Width =1035
                    Height =260
                    FontSize =8
                    TabIndex =44
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Core Dia"
                    ControlSource ="CoreDia"
                    Format ="Fixed"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Core_Dia"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =6240
                    LayoutCachedWidth =2475
                    LayoutCachedHeight =6500
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
                            OverlapFlags =85
                            Left =540
                            Top =6240
                            Width =810
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Core Dia Label"
                            Caption ="Core Dia"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Core_Dia_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =6240
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =6500
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    Left =4995
                    Top =6240
                    Width =1245
                    Height =260
                    FontSize =8
                    TabIndex =45
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Ripcord"
                    ControlSource ="Ripcord"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4995
                    LayoutCachedTop =6240
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =6500
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
                            Left =3960
                            Top =6240
                            Width =1035
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Ripcord Label"
                            Caption ="Ripcord"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Ripcord_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =6240
                            LayoutCachedWidth =4995
                            LayoutCachedHeight =6500
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =7740
                    Top =4440
                    Width =1200
                    Height =260
                    FontSize =8
                    TabIndex =29
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text22"
                    ControlSource ="AramidType1"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =4440
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =4700
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
                            Left =6660
                            Top =4440
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label23"
                            Caption ="Aramid #1"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =4440
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =4700
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =7740
                    Top =4700
                    Width =660
                    Height =260
                    FontSize =8
                    TabIndex =32
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text24"
                    ControlSource ="AramidEnds1"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =4700
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =4960
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
                            OverlapFlags =87
                            TextAlign =1
                            Left =6660
                            Top =4700
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label25"
                            Caption ="Ara Qty #1"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =4700
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =4960
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =7740
                    Top =5040
                    Width =1200
                    Height =260
                    FontSize =8
                    TabIndex =35
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text26"
                    ControlSource ="AramidType2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =5040
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =5300
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
                            Left =6660
                            Top =5040
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label27"
                            Caption ="Aramid #2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =5040
                            LayoutCachedWidth =7680
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =7740
                    Top =5300
                    Width =660
                    Height =260
                    FontSize =8
                    TabIndex =38
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text28"
                    ControlSource ="AramidEnds2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =5300
                    LayoutCachedWidth =8400
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
                            OverlapFlags =87
                            TextAlign =1
                            Left =6660
                            Top =5300
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label29"
                            Caption ="Ara Qty #2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =5300
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =5560
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =2
                    Left =840
                    Top =7155
                    Width =3120
                    Height =260
                    FontSize =8
                    TabIndex =47
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Lay Length"
                    ControlSource ="LayLength"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Lay_Length"
                    AsianLineBreak =0

                    LayoutCachedLeft =840
                    LayoutCachedTop =7155
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =7415
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
                            TextAlign =2
                            Left =840
                            Top =6900
                            Width =3120
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Lay Length Label"
                            Caption ="Lay Length"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Lay_Length_Label"
                            LayoutCachedLeft =840
                            LayoutCachedTop =6900
                            LayoutCachedWidth =3960
                            LayoutCachedHeight =7160
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =2
                    Left =3960
                    Top =7155
                    Width =1560
                    Height =260
                    FontSize =8
                    TabIndex =48
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Helix Factor"
                    ControlSource ="HelixFactor"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Helix_Factor"
                    AsianLineBreak =0

                    LayoutCachedLeft =3960
                    LayoutCachedTop =7155
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =7415
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
                            TextAlign =2
                            Left =3960
                            Top =6900
                            Width =1560
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Helix Factor Label"
                            Caption ="Helix Factor"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Helix_Factor_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =6900
                            LayoutCachedWidth =5520
                            LayoutCachedHeight =7160
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
                    Left =360
                    Top =60
                    Width =2880
                    Height =405
                    FontSize =14
                    FontWeight =600
                    Name ="Label54"
                    Caption ="Cable Construction"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =465
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =2
                    Left =5520
                    Top =7155
                    Width =1560
                    Height =260
                    FontSize =8
                    TabIndex =49
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text8"
                    ControlSource ="Talc"
                    Format ="Yes/No"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5520
                    LayoutCachedTop =7155
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =7415
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
                            TextAlign =2
                            Left =5520
                            Top =6900
                            Width =1560
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label9"
                            Caption ="Talc"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =6900
                            LayoutCachedWidth =7080
                            LayoutCachedHeight =7160
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =7740
                    Top =5640
                    Width =1200
                    Height =260
                    FontSize =8
                    TabIndex =40
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text12"
                    ControlSource ="AramidType3"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =5640
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =5900
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
                            Left =6660
                            Top =5640
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label13"
                            Caption ="Aramid #2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =5640
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =5900
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =7740
                    Top =5900
                    Width =660
                    Height =260
                    FontSize =8
                    TabIndex =43
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text16"
                    ControlSource ="AramidEnds3"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =5900
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =6160
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
                            OverlapFlags =87
                            TextAlign =1
                            Left =6660
                            Top =5900
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label17"
                            Caption ="Ara Qty #2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =5900
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =6160
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
                    OverlapFlags =93
                    TextAlign =1
                    Left =3720
                    Top =2700
                    Width =1016
                    Height =260
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label18"
                    Caption ="OD Tol (+)"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =2700
                    LayoutCachedWidth =4736
                    LayoutCachedHeight =2960
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =3720
                    Top =2960
                    Width =1016
                    Height =260
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label19"
                    Caption ="Height Tol (+)"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =2960
                    LayoutCachedWidth =4736
                    LayoutCachedHeight =3220
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =6600
                    Top =2700
                    Width =1016
                    Height =260
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label20"
                    Caption ="OD Tol (-)"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =6600
                    LayoutCachedTop =2700
                    LayoutCachedWidth =7616
                    LayoutCachedHeight =2960
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =6600
                    Top =2960
                    Width =1016
                    Height =260
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label21"
                    Caption ="Height Tol (-)"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =6600
                    LayoutCachedTop =2960
                    LayoutCachedWidth =7616
                    LayoutCachedHeight =3220
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Left =420
                    Top =7500
                    Width =8504
                    Name ="Line25"
                    LayoutCachedLeft =420
                    LayoutCachedTop =7500
                    LayoutCachedWidth =8924
                    LayoutCachedHeight =7500
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =3480
                    Top =7880
                    Width =1976
                    Height =260
                    FontSize =8
                    TabIndex =51
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text39"
                    ControlSource ="TensileRatingShortTerm_N"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3480
                    LayoutCachedTop =7880
                    LayoutCachedWidth =5456
                    LayoutCachedHeight =8140
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
                            Left =3480
                            Top =7620
                            Width =1976
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label40"
                            Caption ="Tensile - Short Term (N)"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3480
                            LayoutCachedTop =7620
                            LayoutCachedWidth =5456
                            LayoutCachedHeight =7880
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =5460
                    Top =7880
                    Width =1976
                    Height =260
                    FontSize =8
                    TabIndex =52
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text36"
                    ControlSource ="TensileRatingLongTerm_N"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5460
                    LayoutCachedTop =7880
                    LayoutCachedWidth =7436
                    LayoutCachedHeight =8140
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
                            Left =5460
                            Top =7620
                            Width =1976
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label37"
                            Caption ="Tensile - Long Term (N)"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5460
                            LayoutCachedTop =7620
                            LayoutCachedWidth =7436
                            LayoutCachedHeight =7880
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
                Begin ComboBox
                    ColumnHeads = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    ColumnCount =4
                    ListWidth =7200
                    Left =1500
                    Top =7880
                    Width =1976
                    FontSize =8
                    TabIndex =50
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"40\""
                    Name ="Text37"
                    ControlSource ="TemperatureApplication"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT dbo_tblCableTemperatureStandards.TemperatureApplication, dbo_tblCableTemp"
                        "eratureStandards.Operation_C, dbo_tblCableTemperatureStandards.StorageShipping_C"
                        ", dbo_tblCableTemperatureStandards.Installation_C FROM dbo_tblCableTemperatureSt"
                        "andards; "
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =7880
                    LayoutCachedWidth =3476
                    LayoutCachedHeight =8120
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =1500
                            Top =7620
                            Width =1976
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label38"
                            Caption ="Temp"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1500
                            LayoutCachedTop =7620
                            LayoutCachedWidth =3476
                            LayoutCachedHeight =7880
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3960
                    Top =980
                    Width =1500
                    Height =260
                    FontSize =8
                    TabIndex =3
                    BackColor =255
                    ForeColor =-2147483640
                    Name ="txSetup"
                    ControlSource ="StandardOperation"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3960
                    LayoutCachedTop =980
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =1240
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
                            Left =3960
                            Top =720
                            Width =1500
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label4"
                            Caption ="Std Op:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =720
                            LayoutCachedWidth =5460
                            LayoutCachedHeight =980
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =5460
                    Top =980
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txCablePasses"
                    ControlSource ="CablePasses"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5460
                    LayoutCachedTop =980
                    LayoutCachedWidth =6896
                    LayoutCachedHeight =1240
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
                            Left =5460
                            Top =720
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label7"
                            Caption ="Cable Passes"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5460
                            LayoutCachedTop =720
                            LayoutCachedWidth =6900
                            LayoutCachedHeight =980
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =480
                    Top =980
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text15"
                    ControlSource ="EngineeringAssist"
                    Format ="Yes/No"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =480
                    LayoutCachedTop =980
                    LayoutCachedWidth =1916
                    LayoutCachedHeight =1240
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
                            Left =480
                            Top =720
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label16"
                            Caption ="Engineering Assist"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =480
                            LayoutCachedTop =720
                            LayoutCachedWidth =1920
                            LayoutCachedHeight =980
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =1920
                    Top =980
                    Width =2036
                    Height =260
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text47"
                    ControlSource ="EngineeringAssistReason"
                    Format ="Yes/No"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1920
                    LayoutCachedTop =980
                    LayoutCachedWidth =3956
                    LayoutCachedHeight =1240
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
                            Left =1920
                            Top =720
                            Width =2040
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label48"
                            Caption ="Engineering Assist Reason"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1920
                            LayoutCachedTop =720
                            LayoutCachedWidth =3960
                            LayoutCachedHeight =980
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
                    OverlapFlags =85
                    TextAlign =3
                    Left =3900
                    Top =360
                    Width =1230
                    Height =210
                    FontSize =8
                    ForeColor =2366701
                    Name ="lblStdOpExists_Err"
                    Caption ="* Std Op Exists"
                    FontName ="Tahoma"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3900
                    LayoutCachedTop =360
                    LayoutCachedWidth =5130
                    LayoutCachedHeight =570
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =540
                    Top =2240
                    Width =1801
                    Height =260
                    FontSize =8
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Jacket Material"
                    ControlSource ="JacketMaterial"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Jacket_Material"
                    AsianLineBreak =0

                    LayoutCachedLeft =540
                    LayoutCachedTop =2240
                    LayoutCachedWidth =2341
                    LayoutCachedHeight =2500
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
                            Left =540
                            Top =1980
                            Width =1801
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Jacket Material Label"
                            Caption ="Jacket Material"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Jacket_Material_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =1980
                            LayoutCachedWidth =2341
                            LayoutCachedHeight =2240
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =2340
                    Top =2240
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text7"
                    ControlSource ="Additive"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2340
                    LayoutCachedTop =2240
                    LayoutCachedWidth =3776
                    LayoutCachedHeight =2500
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
                            Left =2340
                            Top =1980
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label8"
                            Caption ="Additive"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =2340
                            LayoutCachedTop =1980
                            LayoutCachedWidth =3780
                            LayoutCachedHeight =2240
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =3780
                    Top =2240
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text9"
                    ControlSource ="AdditivePercentage"
                    Format ="Percent"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3780
                    LayoutCachedTop =2240
                    LayoutCachedWidth =5216
                    LayoutCachedHeight =2500
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
                            Left =3780
                            Top =1980
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label10"
                            Caption ="Additive %"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3780
                            LayoutCachedTop =1980
                            LayoutCachedWidth =5220
                            LayoutCachedHeight =2240
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =5220
                    Top =2240
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text38"
                    ControlSource ="Additive2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5220
                    LayoutCachedTop =2240
                    LayoutCachedWidth =6656
                    LayoutCachedHeight =2500
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
                            Left =5220
                            Top =1980
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label43"
                            Caption ="Additive2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5220
                            LayoutCachedTop =1980
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =2240
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =6660
                    Top =2240
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text40"
                    ControlSource ="AdditivePercentage2"
                    Format ="Percent"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6660
                    LayoutCachedTop =2240
                    LayoutCachedWidth =8096
                    LayoutCachedHeight =2500
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
                            Left =6660
                            Top =1980
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label41"
                            Caption ="Additive2 %"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =1980
                            LayoutCachedWidth =8100
                            LayoutCachedHeight =2240
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
                    Left =6720
                    Top =3540
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =18
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text4"
                    ControlSource ="=Round([Nominal OD]*0.0393701,3)"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6720
                    LayoutCachedTop =3540
                    LayoutCachedWidth =8201
                    LayoutCachedHeight =3800
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
                            Left =6720
                            Top =3300
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label5"
                            Caption ="Nominal OD (in)"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6720
                            LayoutCachedTop =3300
                            LayoutCachedWidth =8201
                            LayoutCachedHeight =3560
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
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =4995
                    Top =6480
                    Width =1245
                    Height =260
                    FontSize =8
                    TabIndex =46
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text44"
                    ControlSource ="RipcordQty"
                    Format ="General Number"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4995
                    LayoutCachedTop =6480
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =6740
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
                            Left =3960
                            Top =6480
                            Width =1035
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label45"
                            Caption ="Ripcord Qty"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =6480
                            LayoutCachedWidth =4995
                            LayoutCachedHeight =6740
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
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    ColumnCount =2
                    Left =2220
                    Top =1620
                    Width =2036
                    FontSize =8
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="txFiberApplication"
                    ControlSource ="FiberApplicationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Construction_FiberApplication.FiberApplicationID, Construction_FiberAppli"
                        "cation.FiberApplication FROM Construction_FiberApplication ORDER BY Construction"
                        "_FiberApplication.FiberApplicationID;"
                    ColumnWidths ="0;1440"
                    FontName ="MS Sans Serif"
                    Format ="Yes/No"
                    LayoutCachedLeft =2220
                    LayoutCachedTop =1620
                    LayoutCachedWidth =4256
                    LayoutCachedHeight =1860
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2220
                            Top =1360
                            Width =2040
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label49"
                            Caption ="Fiber Application"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =2220
                            LayoutCachedTop =1360
                            LayoutCachedWidth =4260
                            LayoutCachedHeight =1620
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
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    ColumnCount =2
                    Left =480
                    Top =1620
                    Width =1736
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="txCableShape"
                    ControlSource ="CableShapeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Construction_CableShape.CableShapeID, Construction_CableShape.CableShape "
                        "FROM Construction_CableShape ORDER BY Construction_CableShape.CableShapeID;"
                    ColumnWidths ="0;1440"
                    FontName ="MS Sans Serif"
                    Format ="Yes/No"
                    LayoutCachedLeft =480
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2216
                    LayoutCachedHeight =1860
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =480
                            Top =1360
                            Width =1740
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label47"
                            Caption ="Cable Shape"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =480
                            LayoutCachedTop =1360
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =1620
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
                    Left =7020
                    Top =600
                    Width =666
                    Height =666
                    ForeColor =3547148
                    Name ="btnDeleteRecord"
                    Caption ="Command50"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000d70d0000d70d0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffc7c7c7 ,
                        0x0f0f0f0000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000f0f0fc8c8c8ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff8181810000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000818181ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff6b6b6b0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000006b6b6bffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xff5555550000000000000000000303031818180000000000000000000000000e ,
                        0x0e0e0e0e0e000000000000000000000000181818030303000000000000000000 ,
                        0x555555ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff404040000000000000000000a9a9a9fefefe54 ,
                        0x5454000000000000121212ecececececec111111000000000000545454fefefe ,
                        0xa9a9a9000000000000000000404040ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffff2a2a2a00000000 ,
                        0x0000000000e6e6e6ffffff7979790000000000002e2e2effffffffffff2e2e2e ,
                        0x000000000000797979ffffffe5e5e50000000000000000002a2a2affffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffff131313000000000000030303fbfbfbffffff626262000000000000 ,
                        0x2e2e2effffffffffff2e2e2e000000000000626262fffffff9f9f90101010000 ,
                        0x00000000141414ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffafafa020202000000000000141414ffffff ,
                        0xffffff4b4b4b0000000000002e2e2effffffffffff2e2e2e0000000000004b4b ,
                        0x4bffffffffffff131313000000000000040404fcfcfcffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffe8e8e8000000 ,
                        0x0000000000002c2c2cffffffffffff3333330000000000002e2e2effffffffff ,
                        0xff2e2e2e000000000000333333ffffffffffff2c2c2c000000000000000000e9 ,
                        0xe9e9ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffd2d2d2000000000000000000434343ffffffffffff1c1c1c0000 ,
                        0x000000002e2e2effffffffffff2e2e2e0000000000001c1c1cffffffffffff43 ,
                        0x4343000000000000000000d2d2d2ffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffbdbdbd0000000000000000005a5a ,
                        0x5afffffffefefe0505050000000000002e2e2effffffffffff2e2e2e00000000 ,
                        0x0000060606fefefeffffff5a5a5a000000000000000000bdbdbdffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffa7a7 ,
                        0xa7000000000000000000717171ffffffececec0000000000000000002e2e2eff ,
                        0xffffffffff2e2e2e000000000000000000eeeeeeffffff717171000000000000 ,
                        0x000000a7a7a7ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffff919191000000000000000000898989ffffffd5d5d500 ,
                        0x00000000000000002e2e2effffffffffff2e2e2e000000000000000000d5d5d5 ,
                        0xffffff898989000000000000000000919191ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff7c7c7c00000000000000 ,
                        0x0000a0a0a0ffffffbebebe0000000000000000002e2e2effffffffffff2e2e2e ,
                        0x000000000000000000bebebeffffffa0a0a00000000000000000007c7c7cffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffff666666000000000000000000b7b7b7ffffffa7a7a7000000000000000000 ,
                        0x2e2e2effffffffffff2e2e2e000000000000000000a7a7a7ffffffb7b7b70000 ,
                        0x00000000000000666666ffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffff505050000000000000000000cfcfcfffffff ,
                        0x8f8f8f0000000000000000002e2e2effffffffffff2e2e2e0000000000000000 ,
                        0x008f8f8fffffffcfcfcf000000000000000000505050ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffff3a3a3a000000 ,
                        0x000000000000e7e7e7ffffff7878780000000000000000002e2e2effffffffff ,
                        0xff2e2e2e000000000000000000787878ffffffe6e6e60000000000000000003a ,
                        0x3a3affffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff252525000000000000030303fbfbfbffffff6161610000000000 ,
                        0x000000002e2e2effffffffffff2e2e2e000000000000000000616161fffffffa ,
                        0xfafa010101000000000000252525ffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffff0d0d0d000000000000161616ffff ,
                        0xffffffff4949490000000000000000002e2e2effffffffffff2e2e2e00000000 ,
                        0x0000000000494949ffffffffffff1515150000000000000f0f0fffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f70101 ,
                        0x010000000000002d2d2dffffffffffff3232320000000000000000002e2e2eff ,
                        0xffffffffff2e2e2e000000000000000000323232ffffffffffff2d2d2d000000 ,
                        0x000000020202fafafaffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffe3e3e30000000000000000003d3d3dffffffffffff18181800 ,
                        0x00000000000000002a2a2affffffffffff2a2a2a000000000000000000191919 ,
                        0xffffffffffff3d3d3d000000000000000000e3e3e3ffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffcdcdcd00000000000000000007 ,
                        0x07079999998080800000000000000000000000000202028e8e8e8e8e8e020202 ,
                        0x000000000000000000000000818181999999070707000000000000000000cdcd ,
                        0xcdffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffb7 ,
                        0xb7b7000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000b7b7b7ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffa2a2a2000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000a2a2a2ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffff9f9f9f000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00009f9f9fffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffff8f8f89191918080808080808080808080808080808080808080808080 ,
                        0x8080808080808080808080808080808080808080808080808080808080808080 ,
                        0x8080808080808080808080919191f8f8f8ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffeaea ,
                        0xeadcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdc ,
                        0xdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdcdc ,
                        0xdcdcdceaeaeaffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffbfbfbf03030300000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000040404bfbfbfffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffababab00000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000000000abab ,
                        0xabfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff6 ,
                        0xf6f6191919000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000191919f6f6f6ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffb3b3b3020202000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000030303b3b3b3ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffadadad ,
                        0x1111110000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000111111adadadff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffededed8f8f8f5353534141414141414141412d2d ,
                        0x2d0000000000000000000000000000000000002d2d2d41414141414141414153 ,
                        0x53538f8f8fedededffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffff4a4a4a0000000000000000000000000000000000004a ,
                        0x4a4affffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff7d7d7d02020200000000 ,
                        0x00000000000000000202027d7d7dffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    Picture ="delete.bmp"
                    GridlineColor =10921638

                    LayoutCachedLeft =7020
                    LayoutCachedTop =600
                    LayoutCachedWidth =7686
                    LayoutCachedHeight =1266
                    BorderColor =15916467
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =2
                    ColumnCount =2
                    Left =7080
                    Top =7155
                    Width =1560
                    FontSize =8
                    TabIndex =54
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"20\""
                    Name ="Text51"
                    ControlSource ="JacketStripID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT dbo_tblTightBufferStripType.ID, dbo_tblTightBufferStripType.TightBufferSt"
                        "rip, * FROM dbo_tblTightBufferStripType ORDER BY dbo_tblTightBufferStripType.Tig"
                        "htBufferStrip;"
                    ColumnWidths ="0;2160"
                    FontName ="MS Sans Serif"
                    Format ="Yes/No"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =7155
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =7395
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =2
                            Left =7065
                            Top =6900
                            Width =1575
                            Height =255
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label52"
                            Caption ="Jacket Strip"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =7065
                            LayoutCachedTop =6900
                            LayoutCachedWidth =8640
                            LayoutCachedHeight =7155
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

Private Sub btnDeleteRecord_Click()
Call DeleteARecord
    Me.Parent.SetIsConstructionLoaded (False)
End Sub

Private Sub txNominalWall_AfterUpdate()

Me.txMaxAveWall = Round(1.1 * Me.txNominalWall, 2)
Me.txMinAveWall = Round(0.9 * Me.txNominalWall, 2)
Me.txMinSpotWall = Round(0.8 * Me.txNominalWall, 2)
Me.Requery

End Sub

Public Sub txSetup_AfterUpdate()
Dim Cable As New clsCableConstruction
Dim Exists As Boolean

Dim StdOp As String

StdOp = Me.txSetup

Exists = Cable.SetupExist(StdOp)

'Paint the background of control and show label
Call Me.Parent.PaintExistingRed(Me.txSetup, Me.lblStdOpExists_Err, Exists)

If Exists = False Then

    Me.Parent.newCableCon.StandardOperation = StdOp
End If

End Sub

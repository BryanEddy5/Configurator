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
    ItemSuffix =56
    Left =495
    Top =1320
    Right =9345
    Bottom =11790
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x46846c66eed5e440
    End
    RecordSource ="tblCableConstructions"
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
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            KeepTogether = NotDefault
            Height =10040
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =1721
                    Top =2640
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Nominal OD"
                    ControlSource ="NominalOD"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Nominal_OD"
                    AsianLineBreak =0

                    LayoutCachedLeft =1721
                    LayoutCachedTop =2640
                    LayoutCachedWidth =3202
                    LayoutCachedHeight =2900
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
                            Top =2640
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
                            LayoutCachedTop =2640
                            LayoutCachedWidth =1556
                            LayoutCachedHeight =2900
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
                    OverlapFlags =93
                    TextAlign =1
                    Left =4800
                    Top =2640
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="OD Tol  (+)"
                    ControlSource ="ODTolPlus"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="OD_Tol_____"
                    AsianLineBreak =0

                    LayoutCachedLeft =4800
                    LayoutCachedTop =2640
                    LayoutCachedWidth =6281
                    LayoutCachedHeight =2900
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =7680
                    Top =2640
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="OD Tol  (-)1"
                    ControlSource ="ODTolMinus"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="OD_Tol_____1"
                    AsianLineBreak =0

                    LayoutCachedLeft =7680
                    LayoutCachedTop =2640
                    LayoutCachedWidth =9161
                    LayoutCachedHeight =2900
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =1721
                    Top =2900
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Height"
                    ControlSource ="Height"
                    Format ="Fixed"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1721
                    LayoutCachedTop =2900
                    LayoutCachedWidth =3202
                    LayoutCachedHeight =3160
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
                            Top =2900
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
                            LayoutCachedTop =2900
                            LayoutCachedWidth =1556
                            LayoutCachedHeight =3160
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
                    Left =4800
                    Top =2900
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =14
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Height Tol (+)"
                    ControlSource ="HeightTolPlus"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Height_Tol____"
                    AsianLineBreak =0

                    LayoutCachedLeft =4800
                    LayoutCachedTop =2900
                    LayoutCachedWidth =6281
                    LayoutCachedHeight =3160
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =7680
                    Top =2900
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =15
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Height Tol (-)2"
                    ControlSource ="HeightTolMinus"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Height_Tol____2"
                    AsianLineBreak =0

                    LayoutCachedLeft =7680
                    LayoutCachedTop =2900
                    LayoutCachedWidth =9161
                    LayoutCachedHeight =3160
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =540
                    Top =3498
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =17
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NominalWall"
                    ControlSource ="NominalWall"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =540
                    LayoutCachedTop =3498
                    LayoutCachedWidth =2021
                    LayoutCachedHeight =3758
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
                            Top =3238
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
                            LayoutCachedTop =3238
                            LayoutCachedWidth =2021
                            LayoutCachedHeight =3498
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
                    Left =2021
                    Top =3498
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =18
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MaxAveWall"
                    ControlSource ="MaxAveWall"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2021
                    LayoutCachedTop =3498
                    LayoutCachedWidth =3502
                    LayoutCachedHeight =3758
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
                            Top =3238
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
                            LayoutCachedTop =3238
                            LayoutCachedWidth =3502
                            LayoutCachedHeight =3498
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
                    Left =3502
                    Top =3498
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =19
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MinAveWall"
                    ControlSource ="MinAveWall"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3502
                    LayoutCachedTop =3498
                    LayoutCachedWidth =4983
                    LayoutCachedHeight =3758
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
                            Top =3238
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
                            LayoutCachedTop =3238
                            LayoutCachedWidth =4983
                            LayoutCachedHeight =3498
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
                    Left =4983
                    Top =3498
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =20
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MinSpotWall"
                    ControlSource ="MinSpotWall"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4983
                    LayoutCachedTop =3498
                    LayoutCachedWidth =6464
                    LayoutCachedHeight =3758
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
                            Top =3238
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
                            LayoutCachedTop =3238
                            LayoutCachedWidth =6464
                            LayoutCachedHeight =3498
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
                    Top =3900
                    Width =8504
                    Name ="Line130"
                    LayoutCachedLeft =480
                    LayoutCachedTop =3900
                    LayoutCachedWidth =8984
                    LayoutCachedHeight =3900
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =540
                    Top =2180
                    Width =2041
                    Height =260
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Jacket Material"
                    ControlSource ="JacketMaterial"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Jacket_Material"
                    AsianLineBreak =0

                    LayoutCachedLeft =540
                    LayoutCachedTop =2180
                    LayoutCachedWidth =2581
                    LayoutCachedHeight =2440
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
                            Top =1920
                            Width =2041
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Jacket Material Label"
                            Caption ="Jacket Material"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Jacket_Material_Label"
                            LayoutCachedLeft =540
                            LayoutCachedTop =1920
                            LayoutCachedWidth =2581
                            LayoutCachedHeight =2180
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
                    Top =2520
                    Width =8504
                    Name ="Line0"
                    LayoutCachedLeft =480
                    LayoutCachedTop =2520
                    LayoutCachedWidth =8984
                    LayoutCachedHeight =2520
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Left =420
                    Top =8160
                    Width =8504
                    Name ="Line135"
                    LayoutCachedLeft =420
                    LayoutCachedTop =8160
                    LayoutCachedWidth =8924
                    LayoutCachedHeight =8160
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =1901
                    Top =8280
                    Width =6617
                    Height =860
                    FontSize =8
                    TabIndex =51
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text0"
                    ControlSource ="Instructions"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1901
                    LayoutCachedTop =8280
                    LayoutCachedWidth =8518
                    LayoutCachedHeight =9140
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
                            Top =8280
                            Width =1320
                            Height =540
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label1"
                            Caption ="Construction Instructions"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =420
                            LayoutCachedTop =8280
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =8820
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
                    OverlapFlags =247
                    TextAlign =1
                    Left =1901
                    Top =9120
                    Width =6617
                    Height =860
                    FontSize =8
                    TabIndex =52
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text14"
                    ControlSource ="Instructions2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1901
                    LayoutCachedTop =9120
                    LayoutCachedWidth =8518
                    LayoutCachedHeight =9980
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
                            Top =9120
                            Width =1320
                            Height =540
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label15"
                            Caption ="Construction Instructions"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =420
                            LayoutCachedTop =9120
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =9660
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
                    OverlapFlags =93
                    TextAlign =1
                    Left =1440
                    Top =4020
                    Width =2310
                    Height =260
                    FontSize =8
                    TabIndex =22
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="FRP Dia"
                    ControlSource ="FRP_Dia"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="FRP_Dia"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4020
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =4280
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
                            Top =4020
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
                            LayoutCachedTop =4020
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =4280
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
                    OverlapFlags =247
                    TextAlign =1
                    Left =1440
                    Top =4260
                    Width =2310
                    Height =260
                    FontSize =8
                    TabIndex =26
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM Type"
                    ControlSource ="CM"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_Type"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4260
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =4520
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
                            Top =4275
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
                            LayoutCachedTop =4275
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =4535
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
                    OverlapFlags =85
                    TextAlign =1
                    Left =1440
                    Top =4560
                    Width =2310
                    Height =260
                    FontSize =8
                    TabIndex =29
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM Matrl"
                    ControlSource ="CMMaterial"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_Matrl"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4560
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =4820
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
                            Top =4530
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
                            LayoutCachedTop =4530
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =4790
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
                    OverlapFlags =223
                    TextAlign =2
                    Left =5040
                    Top =4020
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =23
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM OD"
                    ControlSource ="CM_OD"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_OD"
                    AsianLineBreak =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =4020
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =4280
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
                            Top =4020
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
                            LayoutCachedTop =4020
                            LayoutCachedWidth =5115
                            LayoutCachedHeight =4280
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
                    OverlapFlags =93
                    TextAlign =2
                    Left =6240
                    Top =4020
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =24
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM OD Tol +"
                    ControlSource ="CMODTolPlus"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_OD_Tol__"
                    AsianLineBreak =0

                    LayoutCachedLeft =6240
                    LayoutCachedTop =4020
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =4280
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    TextAlign =2
                    Left =7380
                    Top =4020
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =25
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM OD Tol -1"
                    ControlSource ="CMODTolMinus"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_OD_Tol__1"
                    AsianLineBreak =0

                    LayoutCachedLeft =7380
                    LayoutCachedTop =4020
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =4280
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    TextAlign =2
                    Left =5040
                    Top =4260
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =27
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM Wall"
                    ControlSource ="CMWall"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_Wall"
                    AsianLineBreak =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =4260
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =4520
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
                            Top =4275
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
                            LayoutCachedTop =4275
                            LayoutCachedWidth =5115
                            LayoutCachedHeight =4535
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
                    TextAlign =2
                    Left =5040
                    Top =4560
                    Width =1155
                    Height =260
                    FontSize =8
                    TabIndex =30
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CM Min Wall"
                    ControlSource ="CMMinWall"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="CM_Min_Wall"
                    AsianLineBreak =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =4560
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =4820
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
                            Top =4530
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
                            LayoutCachedTop =4530
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =4790
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    Left =1440
                    Top =4920
                    Width =2295
                    Height =260
                    FontSize =8
                    TabIndex =32
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Binder #1"
                    ControlSource ="Binder1"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Binder__1"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4920
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =5180
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
                            Top =4920
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
                            LayoutCachedTop =4920
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =5180
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
                    TextAlign =2
                    Left =5010
                    Top =4920
                    Width =1050
                    Height =260
                    FontSize =8
                    TabIndex =33
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Qty Binder #1"
                    ControlSource ="QtyBinder1"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Qty_Binder__1"
                    AsianLineBreak =0

                    LayoutCachedLeft =5010
                    LayoutCachedTop =4920
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =5180
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
                            Top =4920
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
                            LayoutCachedTop =4920
                            LayoutCachedWidth =5010
                            LayoutCachedHeight =5180
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =1440
                    Top =5160
                    Width =2295
                    Height =260
                    FontSize =8
                    TabIndex =35
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Binder #2"
                    ControlSource ="Binder2"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Binder__2"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =5160
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =5420
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
                            Top =5175
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
                            LayoutCachedTop =5175
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =5435
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
                    TextAlign =2
                    Left =5010
                    Top =5175
                    Width =1050
                    Height =260
                    FontSize =8
                    TabIndex =36
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Qty Binder #2"
                    ControlSource ="QtyBinder2"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Qty_Binder__2"
                    AsianLineBreak =0

                    LayoutCachedLeft =5010
                    LayoutCachedTop =5175
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =5435
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
                            Top =5175
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
                            LayoutCachedTop =5175
                            LayoutCachedWidth =5010
                            LayoutCachedHeight =5435
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1440
                    Top =5520
                    Width =2295
                    Height =260
                    FontSize =8
                    TabIndex =38
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Core Wrap"
                    ControlSource ="CoreWrap"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Core_Wrap"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =5520
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =5780
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
                            Top =5520
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
                            LayoutCachedTop =5520
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =5780
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1440
                    Top =5820
                    Width =2295
                    Height =260
                    FontSize =8
                    TabIndex =40
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Binder #3"
                    ControlSource ="Binder3"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Binder__3"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =5820
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =6080
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
                            Top =5820
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
                            LayoutCachedTop =5820
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =6080
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
                    TextAlign =2
                    Left =5010
                    Top =5820
                    Width =1050
                    Height =260
                    FontSize =8
                    TabIndex =41
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Qty Binder #3"
                    ControlSource ="QtyBinder3"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Qty_Binder__3"
                    AsianLineBreak =0

                    LayoutCachedLeft =5010
                    LayoutCachedTop =5820
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =6080
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
                            Top =5820
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
                            LayoutCachedTop =5820
                            LayoutCachedWidth =5010
                            LayoutCachedHeight =6080
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
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1440
                    Top =6180
                    Width =1035
                    Height =260
                    FontSize =8
                    TabIndex =43
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Core Dia"
                    ControlSource ="CoreDia"
                    Format ="Fixed"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Core_Dia"
                    AsianLineBreak =0

                    LayoutCachedLeft =1440
                    LayoutCachedTop =6180
                    LayoutCachedWidth =2475
                    LayoutCachedHeight =6440
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
                            Top =6180
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
                            LayoutCachedTop =6180
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =6440
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
                    Left =4995
                    Top =6180
                    Width =1245
                    Height =260
                    FontSize =8
                    TabIndex =44
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Ripcord"
                    ControlSource ="Ripcord"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4995
                    LayoutCachedTop =6180
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =6440
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
                            Top =6180
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
                            LayoutCachedTop =6180
                            LayoutCachedWidth =4995
                            LayoutCachedHeight =6440
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =7740
                    Top =4380
                    Width =1200
                    Height =260
                    FontSize =8
                    TabIndex =28
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text22"
                    ControlSource ="AramidType1"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =4380
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =4640
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
                            Top =4380
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label23"
                            Caption ="Aramid #1"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =4380
                            LayoutCachedWidth =7680
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
                    Left =7740
                    Top =4640
                    Width =660
                    Height =260
                    FontSize =8
                    TabIndex =31
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text24"
                    ControlSource ="AramidEnds1"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =4640
                    LayoutCachedWidth =8400
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
                            OverlapFlags =87
                            TextAlign =1
                            Left =6660
                            Top =4640
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label25"
                            Caption ="Ara Qty #1"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =4640
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =4900
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
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =7740
                    Top =4980
                    Width =1200
                    Height =260
                    FontSize =8
                    TabIndex =34
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text26"
                    ControlSource ="AramidType2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =4980
                    LayoutCachedWidth =8940
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
                            TextAlign =1
                            Left =6660
                            Top =4980
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label27"
                            Caption ="Aramid #2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =4980
                            LayoutCachedWidth =7680
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =7740
                    Top =5240
                    Width =660
                    Height =260
                    FontSize =8
                    TabIndex =37
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text28"
                    ControlSource ="AramidEnds2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =5240
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =5500
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
                            Top =5240
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label29"
                            Caption ="Ara Qty #2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =5240
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =5500
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
                    TextAlign =2
                    Left =600
                    Top =7035
                    Width =3120
                    Height =260
                    FontSize =8
                    TabIndex =45
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Lay Length"
                    ControlSource ="LayLength"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Lay_Length"
                    AsianLineBreak =0

                    LayoutCachedLeft =600
                    LayoutCachedTop =7035
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =7295
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
                            Left =600
                            Top =6780
                            Width =3120
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Lay Length Label"
                            Caption ="Lay Length"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Lay_Length_Label"
                            LayoutCachedLeft =600
                            LayoutCachedTop =6780
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =7040
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
                    TextAlign =2
                    Left =3720
                    Top =7035
                    Width =1560
                    Height =260
                    FontSize =8
                    TabIndex =46
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Helix Factor"
                    ControlSource ="HelixFactor"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Helix_Factor"
                    AsianLineBreak =0

                    LayoutCachedLeft =3720
                    LayoutCachedTop =7035
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =7295
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
                            Left =3720
                            Top =6780
                            Width =1560
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Helix Factor Label"
                            Caption ="Helix Factor"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Helix_Factor_Label"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =6780
                            LayoutCachedWidth =5280
                            LayoutCachedHeight =7040
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
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5460
                    Top =980
                    Width =1500
                    Height =260
                    FontSize =8
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text3"
                    ControlSource ="StandardOperation"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5460
                    LayoutCachedTop =980
                    LayoutCachedWidth =6960
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
                            Left =5460
                            Top =720
                            Width =1500
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label4"
                            Caption ="Std Op:"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5460
                            LayoutCachedTop =720
                            LayoutCachedWidth =6960
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =6960
                    Top =980
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text6"
                    ControlSource ="CablePasses"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6960
                    LayoutCachedTop =980
                    LayoutCachedWidth =8396
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
                            Left =6960
                            Top =720
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label7"
                            Caption ="Cable Passes"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6960
                            LayoutCachedTop =720
                            LayoutCachedWidth =8400
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =2581
                    Top =2180
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text7"
                    ControlSource ="Additive"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2581
                    LayoutCachedTop =2180
                    LayoutCachedWidth =4017
                    LayoutCachedHeight =2440
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
                            Left =2581
                            Top =1920
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label8"
                            Caption ="Additive"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =2581
                            LayoutCachedTop =1920
                            LayoutCachedWidth =4021
                            LayoutCachedHeight =2180
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
                    Left =4017
                    Top =2180
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text9"
                    ControlSource ="AdditivePercentage"
                    Format ="Percent"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4017
                    LayoutCachedTop =2180
                    LayoutCachedWidth =5453
                    LayoutCachedHeight =2440
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
                            Left =4017
                            Top =1920
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label10"
                            Caption ="Additive %"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =4017
                            LayoutCachedTop =1920
                            LayoutCachedWidth =5457
                            LayoutCachedHeight =2180
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
                    TextAlign =2
                    Left =5280
                    Top =7035
                    Width =1560
                    Height =260
                    FontSize =8
                    TabIndex =48
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text8"
                    ControlSource ="Talc"
                    Format ="Yes/No"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5280
                    LayoutCachedTop =7035
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =7295
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
                            Left =5280
                            Top =6780
                            Width =1560
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label9"
                            Caption ="Talc"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5280
                            LayoutCachedTop =6780
                            LayoutCachedWidth =6840
                            LayoutCachedHeight =7040
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
                    Left =540
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

                    LayoutCachedLeft =540
                    LayoutCachedTop =980
                    LayoutCachedWidth =1976
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
                            Left =540
                            Top =720
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label16"
                            Caption ="Engineering Assist"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =540
                            LayoutCachedTop =720
                            LayoutCachedWidth =1980
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
                    Left =4020
                    Top =980
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="FirstRun"
                    ControlSource ="FirstRun"
                    Format ="Yes/No"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4020
                    LayoutCachedTop =980
                    LayoutCachedWidth =5456
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
                            Left =4020
                            Top =720
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label11"
                            Caption ="First Run"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =4020
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
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    FELineBreak = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    Left =7740
                    Top =5580
                    Width =1200
                    Height =260
                    FontSize =8
                    TabIndex =39
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text12"
                    ControlSource ="AramidType3"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =5580
                    LayoutCachedWidth =8940
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
                            OverlapFlags =93
                            TextAlign =1
                            Left =6660
                            Top =5580
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label13"
                            Caption ="Aramid #2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =5580
                            LayoutCachedWidth =7680
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
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    Left =7740
                    Top =5840
                    Width =660
                    Height =260
                    FontSize =8
                    TabIndex =42
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text16"
                    ControlSource ="AramidEnds3"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =5840
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =6100
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
                            Top =5840
                            Width =1020
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label17"
                            Caption ="Ara Qty #2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =5840
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =6100
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
                    Top =2640
                    Width =1016
                    Height =260
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label18"
                    Caption ="OD Tol (+)"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =2640
                    LayoutCachedWidth =4736
                    LayoutCachedHeight =2900
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
                    Top =2900
                    Width =1016
                    Height =260
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label19"
                    Caption ="Height Tol (+)"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =2900
                    LayoutCachedWidth =4736
                    LayoutCachedHeight =3160
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
                    Top =2640
                    Width =1016
                    Height =260
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label20"
                    Caption ="OD Tol (-)"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =6600
                    LayoutCachedTop =2640
                    LayoutCachedWidth =7616
                    LayoutCachedHeight =2900
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
                    Top =2900
                    Width =1016
                    Height =260
                    FontSize =8
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label21"
                    Caption ="Height Tol (-)"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =6600
                    LayoutCachedTop =2900
                    LayoutCachedWidth =7616
                    LayoutCachedHeight =3160
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
                    OverlapFlags =215
                    TextAlign =2
                    Left =7140
                    Top =3540
                    Width =1481
                    Height =260
                    FontSize =8
                    TabIndex =21
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text4"
                    ControlSource ="=Round([Nominal OD]*0.0393701,3)"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7140
                    LayoutCachedTop =3540
                    LayoutCachedWidth =8621
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
                            Left =7140
                            Top =3300
                            Width =1481
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label5"
                            Caption ="Nominal OD (in)"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =7140
                            LayoutCachedTop =3300
                            LayoutCachedWidth =8621
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =6600
                    Top =3300
                    Width =426
                    Height =486
                    FontSize =8
                    TabIndex =16
                    ForeColor =-2147483630
                    Name ="btnMils"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada00000dadadadadad0f880adadadadada ,
                        0x0f000dadadadadad0f880adadadadada0f800dadadadadad0f880adadadadada ,
                        0x0f000dadadadadad0f880adadadadada0f800dadadadadad0f880adadadadada ,
                        0x000000000000000d0f8808080808080a0f8808880888080d0fff0fffffffff0a ,
                        0x000000000000000d
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =6600
                    LayoutCachedTop =3300
                    LayoutCachedWidth =7026
                    LayoutCachedHeight =3786
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
                    Enabled = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =2
                    Left =7680
                    Top =380
                    Width =1436
                    Height =260
                    FontSize =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text23"
                    ControlSource ="BaseID"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =7680
                    LayoutCachedTop =380
                    LayoutCachedWidth =9116
                    LayoutCachedHeight =640
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
                            Left =7680
                            Top =120
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label24"
                            Caption ="BaseID"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =7680
                            LayoutCachedTop =120
                            LayoutCachedWidth =9120
                            LayoutCachedHeight =380
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
                    Left =420
                    Top =7380
                    Width =8504
                    Name ="Line25"
                    LayoutCachedLeft =420
                    LayoutCachedTop =7380
                    LayoutCachedWidth =8924
                    LayoutCachedHeight =7380
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    Left =3720
                    Top =7820
                    Width =1976
                    Height =260
                    FontSize =8
                    TabIndex =49
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text39"
                    ControlSource ="TensileRatingShortTerm_N"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3720
                    LayoutCachedTop =7820
                    LayoutCachedWidth =5696
                    LayoutCachedHeight =8080
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
                            Left =3720
                            Top =7560
                            Width =1976
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label40"
                            Caption ="Tensile - Short Term (N)"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =7560
                            LayoutCachedWidth =5696
                            LayoutCachedHeight =7820
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
                    Left =5700
                    Top =7820
                    Width =1976
                    Height =260
                    FontSize =8
                    TabIndex =50
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text36"
                    ControlSource ="TensileRatingLongTerm_N"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5700
                    LayoutCachedTop =7820
                    LayoutCachedWidth =7676
                    LayoutCachedHeight =8080
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
                            Left =5700
                            Top =7560
                            Width =1976
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label37"
                            Caption ="Tensile - Long Term (N)"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5700
                            LayoutCachedTop =7560
                            LayoutCachedWidth =7676
                            LayoutCachedHeight =7820
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
                    Left =6900
                    Top =2180
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text43"
                    ControlSource ="AdditivePercentage2"
                    Format ="Percent"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6900
                    LayoutCachedTop =2180
                    LayoutCachedWidth =8336
                    LayoutCachedHeight =2440
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
                            Left =6900
                            Top =1920
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label44"
                            Caption ="Additive2 %"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6900
                            LayoutCachedTop =1920
                            LayoutCachedWidth =8340
                            LayoutCachedHeight =2180
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
                    Left =5460
                    Top =2180
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text45"
                    ControlSource ="Additive2"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5460
                    LayoutCachedTop =2180
                    LayoutCachedWidth =6896
                    LayoutCachedHeight =2440
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
                            Top =1920
                            Width =1440
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label46"
                            Caption ="Additive2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =5460
                            LayoutCachedTop =1920
                            LayoutCachedWidth =6900
                            LayoutCachedHeight =2180
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
                    Left =1980
                    Top =980
                    Width =2036
                    Height =260
                    FontSize =8
                    TabIndex =53
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text47"
                    ControlSource ="EngineeringAssistReason"
                    Format ="Yes/No"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1980
                    LayoutCachedTop =980
                    LayoutCachedWidth =4016
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
                            Left =1980
                            Top =720
                            Width =2040
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label48"
                            Caption ="Engineering Assist Reason"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1980
                            LayoutCachedTop =720
                            LayoutCachedWidth =4020
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
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    Left =4020
                    Top =120
                    Width =1436
                    Height =260
                    FontSize =8
                    TabIndex =54
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txRevisionHistory"
                    ControlSource ="RevisionHistory"
                    Format ="Yes/No"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4020
                    LayoutCachedTop =120
                    LayoutCachedWidth =5456
                    LayoutCachedHeight =380
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    Locked = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    ColumnCount =4
                    ListWidth =7200
                    Left =1740
                    Top =7820
                    Width =1976
                    FontSize =8
                    TabIndex =47
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

                    LayoutCachedLeft =1740
                    LayoutCachedTop =7820
                    LayoutCachedWidth =3716
                    LayoutCachedHeight =8060
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
                            Left =1740
                            Top =7560
                            Width =1976
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label38"
                            Caption ="Temp"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1740
                            LayoutCachedTop =7560
                            LayoutCachedWidth =3716
                            LayoutCachedHeight =7820
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
                    OverlapFlags =247
                    Left =4995
                    Top =6420
                    Width =1245
                    Height =260
                    FontSize =8
                    TabIndex =55
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text51"
                    ControlSource ="RipcordQty"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =4995
                    LayoutCachedTop =6420
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =6680
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
                            Top =6420
                            Width =1035
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label52"
                            Caption ="Ripcord Qty"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =6420
                            LayoutCachedWidth =4995
                            LayoutCachedHeight =6680
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
                    LimitToList = NotDefault
                    Locked = NotDefault
                    SpecialEffect =2
                    OverlapFlags =95
                    TextAlign =1
                    ColumnCount =2
                    Left =2280
                    Top =1580
                    Width =2036
                    FontSize =8
                    TabIndex =56
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
                    LayoutCachedLeft =2280
                    LayoutCachedTop =1580
                    LayoutCachedWidth =4316
                    LayoutCachedHeight =1820
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
                            Left =2280
                            Top =1320
                            Width =2040
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label49"
                            Caption ="Fiber Application"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =2280
                            LayoutCachedTop =1320
                            LayoutCachedWidth =4320
                            LayoutCachedHeight =1580
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
                    LimitToList = NotDefault
                    Locked = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =1
                    ColumnCount =2
                    Left =540
                    Top =1580
                    Width =1736
                    FontSize =8
                    TabIndex =57
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
                    LayoutCachedLeft =540
                    LayoutCachedTop =1580
                    LayoutCachedWidth =2276
                    LayoutCachedHeight =1820
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
                            Left =540
                            Top =1320
                            Width =1740
                            Height =260
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label47"
                            Caption ="Cable Shape"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =540
                            LayoutCachedTop =1320
                            LayoutCachedWidth =2280
                            LayoutCachedHeight =1580
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
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =87
                    TextAlign =2
                    ColumnCount =2
                    Left =6855
                    Top =7035
                    Width =1560
                    FontSize =8
                    TabIndex =58
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"20\""
                    Name ="Combo53"
                    ControlSource ="JacketStripID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT dbo_tblTightBufferStripType.ID, dbo_tblTightBufferStripType.TightBufferSt"
                        "rip, * FROM dbo_tblTightBufferStripType ORDER BY dbo_tblTightBufferStripType.Tig"
                        "htBufferStrip;"
                    ColumnWidths ="0;2160"
                    FontName ="MS Sans Serif"
                    Format ="Yes/No"

                    LayoutCachedLeft =6855
                    LayoutCachedTop =7035
                    LayoutCachedWidth =8415
                    LayoutCachedHeight =7275
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
                            TextAlign =2
                            Left =6840
                            Top =6780
                            Width =1575
                            Height =255
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label55"
                            Caption ="Jacket Strip"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =6840
                            LayoutCachedTop =6780
                            LayoutCachedWidth =8415
                            LayoutCachedHeight =7035
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

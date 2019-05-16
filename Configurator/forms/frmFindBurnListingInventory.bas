Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =28020
    DatasheetFontHeight =10
    ItemSuffix =77
    Left =3345
    Top =4665
    Right =31830
    Bottom =18765
    DatasheetGridlinesColor =12632256
    Filter ="([ItemNo] like \"*\" OR [Oracle] like \"*\" OR [ListingCompany] like \"*\")  AND"
        " [CablePrint] like \"*OFNP*\""
    OrderBy ="[vBurnListingInventory].[LotNumber] DESC, [vBurnListingInventory].[Item_Cost]"
    RecSrcDt = Begin
        0x57c978a272d2e440
    End
    RecordSource ="vBurnListingInventory"
    Caption ="Item Search"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000ec6d0000e601000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ToggleButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =15790320
            BackThemeColorIndex =6
            BackTint =30.0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
        End
        Begin FormHeader
            Height =1980
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =300
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Oracle"
                    LayoutCachedLeft =300
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =7740
                    Top =1680
                    Width =6240
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text10"
                    Caption ="Print"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =17940
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label15"
                    Caption ="OFNP (FT)"
                    LayoutCachedLeft =17940
                    LayoutCachedTop =1680
                    LayoutCachedWidth =18960
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =25140
                    Top =1680
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="OD (in)"
                    LayoutCachedLeft =25140
                    LayoutCachedTop =1680
                    LayoutCachedWidth =25980
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =2220
                    Top =1680
                    Width =2880
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Catalog Number"
                    LayoutCachedLeft =2220
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =6840
                    Top =1680
                    Width =765
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label24"
                    Caption ="AvailQty"
                    LayoutCachedLeft =6840
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7605
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =19020
                    Top =1680
                    Width =960
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label33"
                    Caption ="OFNR (FT)"
                    LayoutCachedLeft =19020
                    LayoutCachedTop =1680
                    LayoutCachedWidth =19980
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =180
                    Top =420
                    Width =2760
                    Name ="Box35"
                End
                Begin Label
                    OverlapFlags =85
                    Left =26040
                    Top =1680
                    Width =900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label39"
                    Caption ="Status"
                    LayoutCachedLeft =26040
                    LayoutCachedTop =1680
                    LayoutCachedWidth =26940
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =14040
                    Top =1680
                    Width =3840
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="Oracle Description"
                    LayoutCachedLeft =14040
                    LayoutCachedTop =1680
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =20040
                    Top =1680
                    Width =1200
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label43"
                    Caption ="Cost"
                    LayoutCachedLeft =20040
                    LayoutCachedTop =1680
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =21360
                    Top =1680
                    Width =2340
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="Lot Number"
                    LayoutCachedLeft =21360
                    LayoutCachedTop =1680
                    LayoutCachedWidth =23700
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =23760
                    Top =1680
                    Width =1245
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Locator"
                    LayoutCachedLeft =23760
                    LayoutCachedTop =1680
                    LayoutCachedWidth =25005
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1260
                    Width =27720
                    Name ="Line58"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10800
                    Top =780
                    Width =1320
                    TabIndex =1
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =10800
                    LayoutCachedTop =780
                    LayoutCachedWidth =12120
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =300
                    Top =540
                    Width =2400
                    Height =480
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Burn Search"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    Left =5160
                    Top =1680
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label61"
                    Caption ="Listing Co"
                    LayoutCachedLeft =5160
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =9300
                    Top =780
                    Width =1320
                    TabIndex =2
                    Name ="btnSearch"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9300
                    LayoutCachedTop =780
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5580
                    Top =840
                    Width =3600
                    Height =300
                    ColumnOrder =1
                    FontSize =10
                    FontWeight =700
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =840
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =1140
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =5760
                    Top =1320
                    Width =480
                    TabIndex =3
                    Name ="btnSortDescending"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddd444d444dddd0dddd74ddd47ddd707dd ,
                        0xdd44444dddd000dddd74d47ddd70007dddd444dddd00000dddd747dddddd0ddd ,
                        0xdddd4ddddddd0ddddddddddddddd0dddd111111ddddd0dddd11ddd1ddddd0ddd ,
                        0xdd11dddddddd0dddddd11ddddddd0ddddddd11dddddd0dddd1ddd11ddddd0ddd ,
                        0xd111111ddddd0ddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =5760
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =1680
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =5580
                    Top =120
                    Width =3660
                    Height =600
                    FontSize =9
                    Name ="Label65"
                    Caption ="Master Search: Searches all fields displayed with embedded wild card. Ex: PZ0028"
                        "0-01 or ZP002-2.0 or Molex"
                    LayoutCachedLeft =5580
                    LayoutCachedTop =120
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =720
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12300
                    Top =780
                    Width =480
                    TabIndex =4
                    Name ="Refresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dd0000ddd7227ddddd0ffd2d727727dd0f0ffd2727dd727d ,
                        0x0f0ffd227dddd72d0f0ffd2222dddddd0f0ffdddddd2222d0f0ffd27dddd722d ,
                        0x0f0ffd727dd7272d0f0fffd727727d2d0f0ffffd7227dddd0f0fffffd0000ddd ,
                        0x0f0ffffff0f08ddd0f0ffffff008dddd0f000000008ddddd0fffffffdddddddd ,
                        0x00000000dddddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =12300
                    LayoutCachedTop =780
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9300
                    Top =60
                    Width =1320
                    TabIndex =5
                    Name ="Weight"
                    Caption ="Item Weight"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9300
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10800
                    Top =60
                    Width =1320
                    TabIndex =6
                    Name ="btnTensileStrength"
                    Caption ="Tensile"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10800
                    LayoutCachedTop =60
                    LayoutCachedWidth =12120
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12300
                    Top =60
                    Width =1320
                    TabIndex =7
                    Name ="btnBOMExplode"
                    Caption ="Bom Explode"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12300
                    LayoutCachedTop =60
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12900
                    Top =780
                    Width =540
                    TabIndex =9
                    Name ="btnExcelExport"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddd87262ddddddd876262626dddddd6262626262262628 ,
                        0x2626262626fffff662ff6ff2626f62f2268f2f8626fffff6626fff62626f62f2 ,
                        0x2628f82626fffff66268f862626f62f2262fff2626fffff6628f6f82626f62f2 ,
                        0x26ff2ff626fffff662626262626f62f22626262626fffff6d872626262626268 ,
                        0xddddd87626dddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =12900
                    LayoutCachedTop =780
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =27000
                    Top =1680
                    Width =900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label73"
                    Caption ="Passes"
                    LayoutCachedLeft =27000
                    LayoutCachedTop =1680
                    LayoutCachedWidth =27900
                    LayoutCachedHeight =1980
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3180
                    Top =840
                    Width =2160
                    Height =285
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    Name ="cbBurnRating"
                    RowSourceType ="Value List"
                    RowSource ="\"\";OFNP;OFNR;FT4;OF;OFCR;OFCP"
                    FontName ="Times New Roman"
                    OnClick ="[Event Procedure]"
                    Format ="@;\" <pick rating>\""

                    LayoutCachedLeft =3180
                    LayoutCachedTop =840
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1125
                End
                Begin Label
                    OverlapFlags =85
                    Left =3180
                    Top =420
                    Width =2100
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label76"
                    Caption ="Cable Rating Search"
                    LayoutCachedLeft =3180
                    LayoutCachedTop =420
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =720
                End
            End
        End
        Begin Section
            Height =426
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7740
                    Top =60
                    Width =6240
                    Height =300
                    ColumnWidth =1650
                    ColumnOrder =4
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="txCablePrint"
                    ControlSource ="CablePrint"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =60
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =1860
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txOracle"
                    ControlSource ="Oracle"

                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17940
                    Top =60
                    Width =1020
                    Height =300
                    ColumnOrder =5
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="txOFNPBurnLength"
                    ControlSource ="BurnLengthFT_OFNP_UL"

                    LayoutCachedLeft =17940
                    LayoutCachedTop =60
                    LayoutCachedWidth =18960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25140
                    Top =60
                    Width =840
                    Height =300
                    ColumnWidth =2835
                    ColumnOrder =6
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="txNomOD"
                    ControlSource ="NomODInches"

                    LayoutCachedLeft =25140
                    LayoutCachedTop =60
                    LayoutCachedWidth =25980
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2220
                    Top =60
                    Width =2880
                    Height =300
                    ColumnWidth =2280
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txItemNo"
                    ControlSource ="ItemNo"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =60
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6840
                    Top =60
                    Width =780
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="txQty"
                    ControlSource ="AvailQty"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =60
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6300
                    Width =426
                    Height =426
                    Name ="btnFind"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadad000dadada ,
                        0xadadad0bb00dadaddadada0bb00adadaadadada00dadadaddadada0b00dadada ,
                        0xadadad0b00adadaddadada0b00dadadaadadad0bb00dadaddada00d0bb00dada ,
                        0xada0b00d0bb00daddad0b00ad0b00adaada0bb000bb00daddada0bbbbb00dada ,
                        0xadada000000dadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =6300
                    LayoutCachedWidth =6726
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19020
                    Top =60
                    Width =960
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="txOFNRBurnLength"
                    ControlSource ="BurnLengthFT_OFNR_ETL"

                    LayoutCachedLeft =19020
                    LayoutCachedTop =60
                    LayoutCachedWidth =19980
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =26040
                    Top =60
                    Width =900
                    Height =300
                    ColumnWidth =2130
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="txStatus"
                    ControlSource ="Item_Status"

                    LayoutCachedLeft =26040
                    LayoutCachedTop =60
                    LayoutCachedWidth =26940
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14040
                    Top =60
                    Width =3840
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="txOracleDescription"
                    ControlSource ="ItemDesc"

                    LayoutCachedLeft =14040
                    LayoutCachedTop =60
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20040
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="txCost"
                    ControlSource ="Item_Cost"

                    LayoutCachedLeft =20040
                    LayoutCachedTop =60
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21360
                    Top =60
                    Width =2340
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="txLotNumber"
                    ControlSource ="LotNumber"

                    LayoutCachedLeft =21360
                    LayoutCachedTop =60
                    LayoutCachedWidth =23700
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =23760
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txLocator"
                    ControlSource ="Locator"

                    LayoutCachedLeft =23760
                    LayoutCachedTop =60
                    LayoutCachedWidth =25080
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5160
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =12
                    Name ="txListingCompany"
                    ControlSource ="ListingCompanyPrint"

                    LayoutCachedLeft =5160
                    LayoutCachedTop =60
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =27000
                    Top =60
                    Width =900
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =14
                    Name ="txCablePasses"
                    ControlSource ="CablePasses"

                    LayoutCachedLeft =27000
                    LayoutCachedTop =60
                    LayoutCachedWidth =27900
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =1140
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =360
                    Width =1320
                    Name ="btnFiberClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =300
                    LayoutCachedTop =360
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =720
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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
Option Compare Database
Option Explicit
'Key Preview = Yes


Private Sub btnBOMExplode_Click()

Call BomComponentExplode(Me.Oracle)

DoCmd.OpenForm "frmBOMComponentExplode"

Forms!frmBOMComponentExplode!lblItem.Caption = Me.Oracle

End Sub

Private Sub btnCutSheet_Click()
DoCmd.OpenReport "rptCutSheet", acViewPreview, , "Oracle ='" & Me.Oracle & "'", acWindowNormal
End Sub

Private Sub btnExcelExport_Click()


On Error GoTo ErrorHandler

DoCmd.OutputTo acOutputForm, "frmFindBurnListingInventory", acFormatXLS, Environ("USERPROFILE") & "\Desktop\Burn_Listing_Inventory_Results.xls"
Beep
MsgBox "File has successfully been exported to Desktop"



ErrorExit:
Exit Sub

ErrorHandler:
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit
End Sub

Private Sub btnFiberClose_Click()
DoCmd.Close acForm, Me.name

End Sub

Private Sub btnFind_Click()
Call OpenAClient(0, Me.Oracle, Nz(Me.txCablePasses, 2))
End Sub

Private Sub btnRemoveFilter_Click()

Dim strFilter As String
Me.ItemNoSearch.value = Null
Me.ItemNoSearch.Requery
Me.cbBurnRating.value = Null
Me.cbBurnRating.Requery

Me.Filter = ""
Me.OrderBy = ""
Me.FilterOn = False
Me.OrderByOn = False


Me.Filter = strFilter
Me.FilterOn = True
Me.OrderByOn = True

'Call Refresh_Click


End Sub

Private Sub btnSearch_Click()
Dim strFilter As String
strFilter = "([ItemNo] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Oracle] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [ListingCompany] like """ & Me.ItemNoSearch & "*"")"
'strFilter = strFilter & " OR [CablePrint] like """ & Me.ItemNoSearch & "*"")"
strFilter = strFilter & "  AND [CablePrint] like ""*" & Me.cbBurnRating & "*"""
''strFilter = strFilter & " OR [JacketMaterial] like """ & Me.ItemNoSearch & "*"""
''strFilter = strFilter & " OR [StandardOperation] like """ & Me.ItemNoSearch & "*"""
''strFilter = strFilter & " OR [base] like """ & Me.ItemNoSearch & "*"""
''strFilter = strFilter & " OR [Color] like """ & Me.ItemNoSearch & "*"""
''strFilter = strFilter & " OR [PID] like """ & Me.ItemNoSearch & "*"""




Me.Filter = strFilter
Me.FilterOn = True
'Call Refresh_Click
End Sub



Private Sub btnSortDescending_Click()

Me.OrderBy = "[LotNumber] DESC"
Me.OrderByOn = True

'Call Refresh_Click

End Sub

Private Sub btnTensileStrength_Click()
Dim iTensileStrength As Integer

On Error GoTo ErrorHandler:


iTensileStrength = TensileStrengthCalculator(Me.Oracle)

MsgBox Me.Oracle & " max tensile loading is " & iTensileStrength & " Newtons"

ErrorExit:
Exit Sub

ErrorHandler:
If Err.Number = 94 Then
    MsgBox "No strength members found in cable construction.", vbCritical
Else
    MsgBox Err.Number & " " & Err.Description
End If

Resume ErrorExit
End Sub

Private Sub cbBurnRating_Click()
Call btnSearch_Click
End Sub


Private Sub Form_Open(Cancel As Integer)
Call btnSearch_Click
End Sub

Private Sub ItemNoSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Me.btnSearch.SetFocus
        Call btnSearch_Click
    End If

End Sub


Private Sub Refresh_Click()
'Me.Requery
'Me.Refresh
Call btnRemoveFilter_Click
End Sub

Private Sub tglShowObsoleteItems_Click()

Call btnSearch_Click

End Sub

Private Sub Weight_Click()


Dim Weight As Double

Weight = BOM_WeightCalculator(Me.Oracle)

If Weight = 0 Then
    Exit Sub
End If

MsgBox Me.Oracle & " weighs " & Weight & " kg/km"


End Sub

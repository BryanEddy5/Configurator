Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
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
    Width =27960
    DatasheetFontHeight =10
    ItemSuffix =78
    Left =-15578
    Top =-20303
    Right =-14033
    Bottom =-20303
    DatasheetGridlinesColor =12632256
    Filter ="[Fiber2] = 'NONE' AND [Fiber] = 'FBR00388' AND [Base] ='BM006-3.0' AND [COLOR] ="
        " 'BLACK' AND RibbonHighCure = 0  AND ItemNo = 'BM006N302881' AND Customer = 'AFL"
        " GENERIC' AND PID = 'CORNING PRINT'"
    OrderBy ="[CreationDate] DESC"
    RecSrcDt = Begin
        0x9881f1466cd7e440
    End
    RecordSource ="vItemSearch"
    Caption ="Item Search"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
                    Left =180
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Oracle"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =24240
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text8"
                    Caption ="Color"
                    LayoutCachedLeft =24240
                    LayoutCachedTop =1680
                    LayoutCachedWidth =25500
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =22560
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text10"
                    Caption ="Base"
                    LayoutCachedLeft =22560
                    LayoutCachedTop =1680
                    LayoutCachedWidth =24180
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =16740
                    Top =1680
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label15"
                    Caption ="Fiber 1"
                    LayoutCachedLeft =16740
                    LayoutCachedTop =1680
                    LayoutCachedWidth =18120
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =10380
                    Top =1680
                    Width =3060
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Customer"
                    LayoutCachedLeft =10380
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =25560
                    Top =1680
                    Width =600
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label19"
                    Caption ="UOM"
                    LayoutCachedLeft =25560
                    LayoutCachedTop =1680
                    LayoutCachedWidth =26160
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =2100
                    Top =1680
                    Width =3720
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Catalog Number"
                    LayoutCachedLeft =2100
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =7740
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label24"
                    Caption ="Active"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =1680
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =18180
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label33"
                    Caption ="Fiber 2"
                    LayoutCachedLeft =18180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =19440
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
                    Left =13500
                    Top =1680
                    Width =2220
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label39"
                    Caption ="PID"
                    LayoutCachedLeft =13500
                    LayoutCachedTop =1680
                    LayoutCachedWidth =15720
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =15780
                    Top =1680
                    Width =900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="PID Rev"
                    LayoutCachedLeft =15780
                    LayoutCachedTop =1680
                    LayoutCachedWidth =16680
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =19500
                    Top =1680
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label43"
                    Caption ="Jacket"
                    LayoutCachedLeft =19500
                    LayoutCachedTop =1680
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =26220
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label45"
                    Caption ="Passes"
                    LayoutCachedLeft =26220
                    LayoutCachedTop =1680
                    LayoutCachedWidth =26880
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =8460
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="Reason"
                    LayoutCachedLeft =8460
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =93
                    Left =5880
                    Top =1680
                    Width =1245
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Creation Date"
                    LayoutCachedLeft =5880
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1260
                    Width =27720
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =27900
                    LayoutCachedHeight =1260
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8520
                    Top =780
                    Width =1320
                    TabIndex =1
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =8520
                    LayoutCachedTop =780
                    LayoutCachedWidth =9840
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
                    Caption ="Item Search"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    Left =21420
                    Top =1680
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label61"
                    Caption ="Std Op"
                    LayoutCachedLeft =21420
                    LayoutCachedTop =1680
                    LayoutCachedWidth =22500
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7020
                    Top =780
                    Width =1320
                    TabIndex =2
                    Name ="btnSearch"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =780
                    LayoutCachedWidth =8340
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
                    Left =3180
                    Top =780
                    Width =3600
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    Name ="ItemNoSearch"
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3180
                    LayoutCachedTop =780
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =6360
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

                    LayoutCachedLeft =6360
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1680
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =3180
                    Top =60
                    Width =3660
                    Height =600
                    FontSize =9
                    Name ="Label65"
                    Caption ="Master Search: Searches all fields displayed with embedded wild card. Ex: PZ0028"
                        "0-01 or ZP002-2.0 or Molex"
                    LayoutCachedLeft =3180
                    LayoutCachedTop =60
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =660
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10020
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

                    LayoutCachedLeft =10020
                    LayoutCachedTop =780
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ToggleButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =11340
                    Top =780
                    Width =1500
                    Height =360
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =5
                    ForeColor =3547148
                    Name ="tglShowObsoleteItems"
                    Caption ="Show Obsolete"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =11340
                    LayoutCachedTop =780
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =1140
                    BackColor =16513198
                    OldBorderStyle =6
                    BorderLineStyle =5
                    BorderColor =15916467
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7020
                    Top =60
                    Width =1320
                    TabIndex =6
                    Name ="Weight"
                    Caption ="Item Weight"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =60
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15180
                    Top =60
                    Width =1320
                    TabIndex =7
                    Name ="btnCutSheet"
                    Caption ="Cut Sheet"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =15180
                    LayoutCachedTop =60
                    LayoutCachedWidth =16500
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8520
                    Top =60
                    Width =1320
                    TabIndex =8
                    Name ="btnTensileStrength"
                    Caption ="Tensile"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =60
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10020
                    Top =60
                    Width =1320
                    TabIndex =9
                    Name ="btnBOMExplode"
                    Caption ="Bom Explode"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10020
                    LayoutCachedTop =60
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10620
                    Top =780
                    Width =540
                    TabIndex =10
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

                    LayoutCachedLeft =10620
                    LayoutCachedTop =780
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =20640
                    Top =1680
                    Width =720
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label73"
                    Caption ="OD"
                    LayoutCachedLeft =20640
                    LayoutCachedTop =1680
                    LayoutCachedWidth =21360
                    LayoutCachedHeight =1980
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =13200
                    Top =840
                    Width =3180
                    Height =285
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbOracleStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblDesignCodeJacket.JacketDescription_SalesForce, tblDesignTypes"
                        ".SalesForceDescription FROM tblDesignCodeJacket INNER JOIN (tblCableConstruction"
                        "References INNER JOIN tblDesignTypes ON tblCableConstructionReferences.DesignTyp"
                        "eID = tblDesignTypes.DesignTypeID) ON (tblCableConstructionReferences.CableFamil"
                        "y = tblDesignCodeJacket.CableLetter) AND (tblDesignCodeJacket.JacketLetter = tbl"
                        "CableConstructionReferences.CableLevel1) WHERE (((tblDesignTypes.SalesForceDescr"
                        "iption)=[Forms]![frmFindConstruction]![cbFamilyDesc]));"

                    LayoutCachedLeft =13200
                    LayoutCachedTop =840
                    LayoutCachedWidth =16380
                    LayoutCachedHeight =1125
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =13200
                    Top =540
                    Width =1230
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label96"
                    Caption ="Oracle Status"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =540
                    LayoutCachedWidth =14430
                    LayoutCachedHeight =840
                End
                Begin ComboBox
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =16500
                    Top =840
                    Width =4695
                    Height =285
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =12
                    Name ="cbFamilyDesc"
                    RowSourceType ="Value List"
                    RowSource ="Active;Obsolete;Discontd"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =16500
                    LayoutCachedTop =840
                    LayoutCachedWidth =21195
                    LayoutCachedHeight =1125
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =16500
                    Top =540
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label95"
                    Caption ="Family Desc"
                    LayoutCachedLeft =16500
                    LayoutCachedTop =540
                    LayoutCachedWidth =20280
                    LayoutCachedHeight =840
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13140
                    Top =60
                    Width =1320
                    TabIndex =13
                    Name ="btnWhereUsed"
                    Caption ="Where Used"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =13140
                    LayoutCachedTop =60
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11520
                    Top =60
                    TabIndex =14
                    Name ="btnBomCostExplode"
                    Caption ="Bom Cost Explode"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11520
                    LayoutCachedTop =60
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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
                    Left =24240
                    Top =60
                    Width =1260
                    Height =300
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =500
                    TabIndex =14
                    Name ="txColor"
                    ControlSource ="Color"

                    LayoutCachedLeft =24240
                    LayoutCachedTop =60
                    LayoutCachedWidth =25500
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22560
                    Top =60
                    Width =1620
                    Height =300
                    ColumnWidth =1650
                    ColumnOrder =4
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="txBase"
                    ControlSource ="Base"

                    LayoutCachedLeft =22560
                    LayoutCachedTop =60
                    LayoutCachedWidth =24180
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1860
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txOracle"
                    ControlSource ="Oracle"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16740
                    Top =60
                    Width =1380
                    Height =300
                    ColumnOrder =5
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="txFiber1"
                    ControlSource ="Fiber"

                    LayoutCachedLeft =16740
                    LayoutCachedTop =60
                    LayoutCachedWidth =18120
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10380
                    Top =60
                    Width =3060
                    Height =300
                    ColumnWidth =2835
                    ColumnOrder =6
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="txCustomer"
                    ControlSource ="Customer"

                    LayoutCachedLeft =10380
                    LayoutCachedTop =60
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25560
                    Top =60
                    Width =720
                    Height =300
                    ColumnOrder =7
                    FontSize =9
                    FontWeight =500
                    TabIndex =15
                    Name ="txUOM"
                    ControlSource ="PrintSpacing"

                    LayoutCachedLeft =25560
                    LayoutCachedTop =60
                    LayoutCachedWidth =26280
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =60
                    Width =3720
                    Height =300
                    ColumnWidth =2280
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txItemNo"
                    ControlSource ="ItemNo"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7740
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="txActive"
                    ControlSource ="Active"
                    Format ="Yes/No"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =60
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7260
                    Width =426
                    Height =426
                    Name ="btnFind"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000012000000120000000100180000000000f0030000110b0000110b0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffff6262620000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000636363ffffffffffffffff0000007070 ,
                        0x70bebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebe707070000000ffffffffffbebebe000000bebebeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffbebebe ,
                        0x000000ffffffffffbebebe000000bebebeffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffbebebe000000ffffffbeff ,
                        0xbebebe000000bebebeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffbebebe000000ffffff00ffbebebe000000bebe ,
                        0xbeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffbebebe000000ffffffffffbebebe000000bebebeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffbebebe ,
                        0x000000ffffff00ffbebebe000000bebebeffffffffffffffffffffffffffffff ,
                        0xffffff979797ffffffffffffffffffffffffffffffffffff0f0f0fffffffffff ,
                        0xbebebe000000bebebeffffffffffffffffffffffffffffff9696960000007171 ,
                        0x71ffffffffffffffffffffffffffffffffffffffffffffffbebebe000000bebe ,
                        0xbeffffffffffffffffffffffffffffffffffff707070000000717171ffffffff ,
                        0xffffffffffffffffffffffffffffffffbebebe000000bebebeffffffffffffff ,
                        0xffffffffffffffffffffffffffff707070000000717171ffffffffffffffffff ,
                        0x7f7f7fffffffffffbebebe000000bebebeffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff707070000000717171ffffffbebebe000000ffffffffff ,
                        0xbebebe000000bebebeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff707070000000717171bebebe000000ffffff3dffbebebe000000bebe ,
                        0xbeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff70 ,
                        0x70700000003d3d3d000000ffffff00ffffffff000000707070bebebebebebebe ,
                        0xbebebebebebebebeffffffffffffffffffffffffbebebebebebe3c3c3c000000 ,
                        0x000000ffffffbeffffffff626262000000000000000000000000000000000000 ,
                        0x101010ffffffffffff7f7f7f000000000000000000000000000000ffffffffff ,
                        0xffffffffffffffffffbebebebebebebebebebebebebebebeffffffffffffffff ,
                        0xffffffffbebebebebebebebebebebebeffffffffffffbeff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="OpenIcon.bmp"

                    LayoutCachedLeft =7260
                    LayoutCachedWidth =7686
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
                    Left =26340
                    Top =60
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =16
                    Name ="txLocation"
                    ControlSource ="Location"

                    LayoutCachedLeft =26340
                    LayoutCachedTop =60
                    LayoutCachedWidth =26700
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18180
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="txFiber2"
                    ControlSource ="Fiber2"

                    LayoutCachedLeft =18180
                    LayoutCachedTop =60
                    LayoutCachedWidth =19440
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13500
                    Top =60
                    Width =2220
                    Height =300
                    ColumnWidth =2130
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="txPID"
                    ControlSource ="PID"

                    LayoutCachedLeft =13500
                    LayoutCachedTop =60
                    LayoutCachedWidth =15720
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15780
                    Top =60
                    Width =900
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="txCustomerRevision"
                    ControlSource ="CustomerRev"

                    LayoutCachedLeft =15780
                    LayoutCachedTop =60
                    LayoutCachedWidth =16680
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19500
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="txJacket"
                    ControlSource ="JacketMaterial"

                    LayoutCachedLeft =19500
                    LayoutCachedTop =60
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8460
                    Top =60
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="txReason"
                    ControlSource ="Reason"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =60
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5880
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txCreationDate"
                    ControlSource ="CreationDate"
                    Format ="Short Date"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21420
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =12
                    Name ="txSetup"
                    ControlSource ="StandardOperation"

                    LayoutCachedLeft =21420
                    LayoutCachedTop =60
                    LayoutCachedWidth =22500
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20640
                    Top =60
                    Width =720
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =17
                    Name ="Text74"
                    ControlSource ="NominalOD"

                    LayoutCachedLeft =20640
                    LayoutCachedTop =60
                    LayoutCachedWidth =21360
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =26760
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =18
                    Name ="txOracleStatus"
                    ControlSource ="OracleStatus"

                    LayoutCachedLeft =26760
                    LayoutCachedTop =60
                    LayoutCachedWidth =27840
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =1200
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
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


Private Sub btnBomCostExplode_Click()
DoCmd.SetWarnings False
Call BomCostExplode(Me.Oracle)

DoCmd.OpenForm "frmBOMCostExplode"

Forms!frmBOMCostExplode!lblItem.Caption = Me.Oracle
Forms!frmBOMCostExplode!lblDesignCode.Caption = Me.ItemNo
End Sub

Private Sub btnBOMExplode_Click()
DoCmd.SetWarnings False
Call BomComponentExplode(Me.Oracle)

DoCmd.OpenForm "frmBOMComponentExplode"

Forms!frmBOMComponentExplode!lblItem.Caption = Me.Oracle

End Sub

Private Sub btnCutSheet_Click()
OpenCutSheet (Me.Oracle)
End Sub

Private Sub btnExcelExport_Click()


On Error GoTo ErrorHandler

DoCmd.OutputTo acOutputForm, "frmFindItem", acFormatXLS, Environ("USERPROFILE") & "\Desktop\ItemSearchResults.xls"
Beep
MsgBox "File has successfully been exported to Desktop"



ErrorExit:
Exit Sub

ErrorHandler:
If Err.Number = 2302 Then
    MsgBox "Close the excel spreadsheet and try again", vbCritical
Else
    MsgBox Err.Number & " " & Err.Description
End If
Resume ErrorExit
End Sub

Private Sub btnFiberClose_Click()
DoCmd.Close acForm, Me.name

End Sub

Private Sub btnFind_Click()
Call OpenAClient(0, Me.Oracle, Nz(Me.Location, 2))
End Sub

Private Sub btnRemoveFilter_Click()

Me.ItemNoSearch.value = Null
Me.ItemNoSearch.Requery

'If Me.tglShowObsoleteItems <> 0 Then
Me.Filter = ""
Me.FilterOn = False
'Else
'strFilter = " OracleStatus <> 'Obsolete' AND OracleStatus <> 'Discontd'"
'End If

'Me.filter = strFilter
'Me.FilterOn = True
'Me.OrderByOn = True

'Call Refresh_Click

End Sub





Private Sub btnSearch_Click()



Dim strFilter As String
strFilter = "([ItemNo] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Oracle] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Customer] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Fiber] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Fiber2] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [JacketMaterial] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [StandardOperation] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [base] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Color] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [PID] like """ & Me.ItemNoSearch & "*"")"

'If Me.tglShowObsoleteItems <> 0 Then
'Else
'strFilter = strFilter & " AND [OracleStatus] = ""active"""
'End If



Me.Filter = strFilter
Me.FilterOn = True
'Call Refresh_Click
End Sub





Private Sub btnSortDescending_Click()

Me.OrderBy = "[CreationDate] DESC"
Me.OrderByOn = True

'Call Refresh_Click

End Sub

Private Sub btnTensileStrength_Click()
Dim iTensileStrength As Integer
DoCmd.SetWarnings False
On Error GoTo ErrorHandler:


iTensileStrength = TensileStrengthCalculator(Me.Oracle)

MsgBox Me.Oracle & " max tensile loading is " & iTensileStrength & " Newtons / " & Round(iTensileStrength * 0.2248, 0) & " lbs-f " & vbCrLf & vbCrLf _
        & "Disclaimer: This is the theoretical maximum calculated load.  Not real-world applicable."
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

Private Sub btnWhereUsed_Click()
Call WhereUsed(Me.Oracle)
DoCmd.OpenForm "frmWhereUsed"
Forms!frmWhereUsed!lblItem.Caption = Me.Oracle

End Sub

Private Sub cbFamilyDesc_AfterUpdate()
Call btnSearch_Click
End Sub



Private Sub Form_Close()
'Call CloseItemSearch

End Sub

Private Sub Form_Load()
DoCmd.SetWarnings False
'Call btnSearch_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'Call btnSearch_Click
'Call btnRemoveFilter_Click
Call UserCheck
If sUser = 1 Then
btnCutSheet.Visible = True
Else
btnCutSheet.Visible = False
End If
    
End Sub

Private Sub ItemNoSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Me.btnSearch.SetFocus
        Call btnSearch_Click
    End If

End Sub


Private Sub Refresh_Click()
Me.Requery
'Me.Refresh
End Sub

Private Sub tglShowObsoleteItems_Click()

Call btnSearch_Click

End Sub

Private Sub txItemNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call btnBomCostExplode_Click
End If
End Sub



Private Sub Weight_Click()


Dim Weight As Double

Weight = BOM_WeightCalculator(Me.Oracle)

If Weight = 0 Then
    Exit Sub
End If

MsgBox Me.Oracle & " weighs " & vbCrLf & Weight & " kg/km" & vbCrLf & RoundRanges(Weight * 0.671969) & " lb/kft"


End Sub

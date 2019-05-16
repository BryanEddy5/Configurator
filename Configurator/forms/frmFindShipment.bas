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
    Width =27960
    DatasheetFontHeight =10
    ItemSuffix =76
    Left =930
    Top =4875
    Right =19050
    Bottom =10920
    DatasheetGridlinesColor =12632256
    Filter ="((([ITEM_NUMBER] Like \"*\")) Or (([ItemNo] Like \"*\"))) And (vCable_Shipment_S"
        "ales_History.BILL_TO_NAME Like \"dmsi*\")"
    OrderBy ="[CreationDate] DESC"
    RecSrcDt = Begin
        0x2c6029650fe0e440
    End
    RecordSource ="vCable_Shipment_Sales_History"
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
                    Left =23520
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text8"
                    Caption ="Color"
                    LayoutCachedLeft =23520
                    LayoutCachedTop =1680
                    LayoutCachedWidth =24780
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =21840
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text10"
                    Caption ="Base"
                    LayoutCachedLeft =21840
                    LayoutCachedTop =1680
                    LayoutCachedWidth =23460
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =16080
                    Top =1680
                    Width =1395
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label15"
                    Caption ="Price/ Meter"
                    LayoutCachedLeft =16080
                    LayoutCachedTop =1680
                    LayoutCachedWidth =17475
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =10260
                    Top =1680
                    Width =3060
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Customer"
                    LayoutCachedLeft =10260
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =24840
                    Top =1680
                    Width =600
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label19"
                    Caption ="UOM"
                    LayoutCachedLeft =24840
                    LayoutCachedTop =1680
                    LayoutCachedWidth =25440
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
                    Left =7620
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label24"
                    Caption ="UOM"
                    LayoutCachedLeft =7620
                    LayoutCachedTop =1680
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =17520
                    Top =1680
                    Width =4260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label33"
                    Caption ="Cable Description"
                    LayoutCachedLeft =17520
                    LayoutCachedTop =1680
                    LayoutCachedWidth =21780
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
                    Left =13380
                    Top =1680
                    Width =1185
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label39"
                    Caption ="Ship to State"
                    LayoutCachedLeft =13380
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14565
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =14640
                    Top =1680
                    Width =1395
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="Ship to Country"
                    LayoutCachedLeft =14640
                    LayoutCachedTop =1680
                    LayoutCachedWidth =16035
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =25500
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label45"
                    Caption ="Passes"
                    LayoutCachedLeft =25500
                    LayoutCachedTop =1680
                    LayoutCachedWidth =26160
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =8340
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="Revenue"
                    LayoutCachedLeft =8340
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =93
                    Left =6420
                    Top =1680
                    Width =1140
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Qty"
                    LayoutCachedLeft =6420
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7560
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
                    ColumnOrder =2
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
                    ColumnOrder =1
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
                    Left =11760
                    Top =60
                    Width =1320
                    TabIndex =7
                    Name ="btnCutSheet"
                    Caption ="Cut Sheet"

                    LayoutCachedLeft =11760
                    LayoutCachedTop =60
                    LayoutCachedWidth =13080
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
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =13200
                    Top =840
                    Width =4800
                    Height =285
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"10\";\"140\""
                    Name ="cbOracleStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblDesignTypes.SalesForceDescription FROM tblDesignCodeJacket IN"
                        "NER JOIN (tblCableConstructionReferences INNER JOIN tblDesignTypes ON tblCableCo"
                        "nstructionReferences.DesignTypeID = tblDesignTypes.DesignTypeID) ON (tblDesignCo"
                        "deJacket.JacketLetter = tblCableConstructionReferences.CableLevel1) AND (tblDesi"
                        "gnCodeJacket.CableLetter = tblCableConstructionReferences.CableFamily) ORDER BY "
                        "tblDesignTypes.SalesForceDescription;"

                    LayoutCachedLeft =13200
                    LayoutCachedTop =840
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =1125
                End
                Begin Label
                    OverlapFlags =87
                    Left =13200
                    Top =540
                    Width =1230
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label96"
                    Caption ="Family Desc"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =540
                    LayoutCachedWidth =14430
                    LayoutCachedHeight =840
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
                    Left =23520
                    Top =60
                    Width =1260
                    Height =300
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="txColor"

                    LayoutCachedLeft =23520
                    LayoutCachedTop =60
                    LayoutCachedWidth =24780
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21840
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =12
                    Name ="txBase"

                    LayoutCachedLeft =21840
                    LayoutCachedTop =60
                    LayoutCachedWidth =23460
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
                    Name ="txItem_Number"
                    ControlSource ="ITEM_NUMBER"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16080
                    Top =60
                    Width =1380
                    Height =300
                    ColumnOrder =5
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="txFiber1"
                    ControlSource ="Price_per_meter"

                    LayoutCachedLeft =16080
                    LayoutCachedTop =60
                    LayoutCachedWidth =17460
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10260
                    Top =60
                    Width =3060
                    Height =300
                    ColumnWidth =2835
                    ColumnOrder =6
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="txCustomer"
                    ControlSource ="BILL_TO_NAME"

                    LayoutCachedLeft =10260
                    LayoutCachedTop =60
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =24840
                    Top =60
                    Width =720
                    Height =300
                    ColumnOrder =7
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="txUOM"

                    LayoutCachedLeft =24840
                    LayoutCachedTop =60
                    LayoutCachedWidth =25560
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
                    Left =7620
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="txActive"
                    ControlSource ="PRIMARY_UOM_CODE"
                    Format ="Yes/No"

                    LayoutCachedLeft =7620
                    LayoutCachedTop =60
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5880
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

                    LayoutCachedLeft =5880
                    LayoutCachedWidth =6306
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
                    Left =25620
                    Top =60
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =14
                    Name ="txLocation"
                    ControlSource ="CablePasses"

                    LayoutCachedLeft =25620
                    LayoutCachedTop =60
                    LayoutCachedWidth =25980
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17520
                    Top =60
                    Width =4260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="txFiber2"
                    ControlSource ="SalesForceDescription"

                    LayoutCachedLeft =17520
                    LayoutCachedTop =60
                    LayoutCachedWidth =21780
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13380
                    Top =60
                    Width =1200
                    Height =300
                    ColumnWidth =2130
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="txPID"
                    ControlSource ="SHIP_TO_STATE"

                    LayoutCachedLeft =13380
                    LayoutCachedTop =60
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =60
                    Width =1380
                    Height =270
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="txCustomerRevision"
                    ControlSource ="SHIP_TO_COUNTRY"

                    LayoutCachedLeft =14640
                    LayoutCachedTop =60
                    LayoutCachedWidth =16020
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8340
                    Top =60
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txReason"
                    ControlSource ="REVENUE"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =60
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6420
                    Top =60
                    Width =1155
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="txCreationDate"
                    ControlSource ="PRIMARY_QUANTITY"

                    LayoutCachedLeft =6420
                    LayoutCachedTop =60
                    LayoutCachedWidth =7575
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =26040
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =15
                    Name ="txOracleStatus"

                    LayoutCachedLeft =26040
                    LayoutCachedTop =60
                    LayoutCachedWidth =27120
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
                    Name ="btnClose"
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





Private Sub btnExcelExport_Click()


On Error GoTo ErrorHandler

DoCmd.OutputTo acOutputForm, "frmFindItem", acFormatXLS, Environ("USERPROFILE") & "\Desktop\Shipment_Search_Results.xls"
Beep
MsgBox "File has successfully been exported to Desktop"



ErrorExit:
Exit Sub

ErrorHandler:
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit
End Sub

Private Sub btnClose_Click()
DoCmd.Close acForm, Me.name

End Sub

Private Sub btnFind_Click()
Call OpenAClient(0, Me.ITEM_NUMBER, Nz(Me.CablePasses, 2))
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
strFilter = "[ITEM_NUMBER] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [ItemNo] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Bill_To_Name] like """ & Me.ItemNoSearch & "*"""
'strFilter = strFilter & " OR [Fiber] like """ & Me.ItemNoSearch & "*"""
'strFilter = strFilter & " OR [Fiber2] like """ & Me.ItemNoSearch & "*"""
'strFilter = strFilter & " OR [JacketMaterial] like """ & Me.ItemNoSearch & "*"""
'strFilter = strFilter & " OR [StandardOperation] like """ & Me.ItemNoSearch & "*"""
'strFilter = strFilter & " OR [base] like """ & Me.ItemNoSearch & "*"""
'strFilter = strFilter & " OR [Color] like """ & Me.ItemNoSearch & "*"""
'strFilter = strFilter & " OR [PID] like """ & Me.ItemNoSearch & "*"""

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


Private Sub cbFamilyDesc_AfterUpdate()
Call btnSearch_Click
End Sub

Private Sub Form_Load()
'Call btnSearch_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
Call btnSearch_Click

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

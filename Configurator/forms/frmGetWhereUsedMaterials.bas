Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
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
    Width =25380
    DatasheetFontHeight =10
    ItemSuffix =122
    Left =2235
    Top =5610
    Right =27615
    Bottom =13635
    DatasheetGridlinesColor =12632256
    Filter ="([tblGetWhereUsedMaterial].[Material]=\"ARA00124\")"
    OrderBy ="[tblGetWhereUsedMaterial].[Customer], [tblGetWhereUsedMaterial].[AssemblyItemNum"
        "ber], [tblGetWhereUsedMaterial].[Description] DESC, [tblGetWhereUsedMaterial].[T"
        "ransactionDate] DESC, [tblGetWhereUsedMaterial].[Shipped] DESC, [tblGetWhereUsed"
        "Material].[Item_Status], [tblGetWhereUsedMaterial].[CountPerUOM], [tblGetWhereUs"
        "edMaterial].[MaterialQuantity_PerPrimaryUOM], [tblGetWhereUsedMaterial].[Materia"
        "l]"
    RecSrcDt = Begin
        0x78f0d8556ff4e440
    End
    RecordSource ="tblGetWhereUsedMaterial"
    Caption ="Item Search"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
        Begin FormHeader
            Height =1980
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Material"
                    LayoutCachedLeft =780
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =7740
                    Top =1680
                    Width =1230
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Qty Per UOM"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =1680
                    LayoutCachedWidth =8970
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =10320
                    Top =1680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="UOM"
                    LayoutCachedLeft =10320
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Top =1380
                    Width =24900
                    Name ="Line58"
                    LayoutCachedTop =1380
                    LayoutCachedWidth =24900
                    LayoutCachedHeight =1380
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =3000
                    Top =300
                    Width =4740
                    Name ="Box35"
                    LayoutCachedLeft =3000
                    LayoutCachedTop =300
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =3090
                    Top =420
                    Width =4545
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Material Where Used"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3090
                    LayoutCachedTop =420
                    LayoutCachedWidth =7635
                    LayoutCachedHeight =915
                End
                Begin Label
                    OverlapFlags =85
                    Left =11160
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label81"
                    Caption ="UOM"
                    LayoutCachedLeft =11160
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =9000
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label83"
                    Caption ="Qty Per Order"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =12480
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label87"
                    Caption ="FG Status"
                    LayoutCachedLeft =12480
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =15120
                    Top =1680
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label89"
                    Caption ="Customer"
                    LayoutCachedLeft =15120
                    LayoutCachedTop =1680
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8220
                    Top =840
                    Width =1320
                    TabIndex =3
                    Name ="btnOpenInputs"
                    Caption ="Inputs"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8220
                    LayoutCachedTop =840
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =1200
                    BorderWidth =1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =9660
                    Top =840
                    Width =1320
                    TabIndex =4
                    ForeColor =0
                    Name ="btnRun"
                    Caption ="FG Run"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =840
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =1200
                    BorderWidth =1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13980
                    Top =840
                    Width =540
                    TabIndex =7
                    Name ="btnExcelExport"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dd8444ddddddddddddd8744ddddddddd44444444dddddddd ,
                        0x44444444ddddddddddd8744d726ddddddd8444d6262ddddddddddd6262626262 ,
                        0xdddd26f6f62ffff6dddd62f8f26262f2dddd262f262ffff6dddd626f626262f2 ,
                        0xdddd26f8f62ffff6dddd62f2f26262f2dddd2626262ffff6dddd626262626262 ,
                        0xddddddd8762ddddd000000000000000000000000000000000000000000000000 ,
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

                    LayoutCachedLeft =13980
                    LayoutCachedTop =840
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =1200
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9638
                    Top =180
                    Width =1701
                    Height =315
                    ColumnOrder =0
                    FontSize =11
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tb_FunctionTimer"
                    GridlineColor =10921638
                    VerticalAnchor =1

                    LayoutCachedLeft =9638
                    LayoutCachedTop =180
                    LayoutCachedWidth =11339
                    LayoutCachedHeight =495
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8100
                            Top =180
                            Width =1470
                            Height =315
                            FontSize =11
                            BorderColor =8355711
                            Name ="Label18"
                            Caption ="Run Timer"
                            GridlineColor =10921638
                            VerticalAnchor =1
                            LayoutCachedLeft =8100
                            LayoutCachedTop =180
                            LayoutCachedWidth =9570
                            LayoutCachedHeight =495
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =2460
                    Top =1680
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label96"
                    Caption ="Oracle Item"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =17100
                    Top =1680
                    Width =1230
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label98"
                    Caption ="Order #"
                    LayoutCachedLeft =17100
                    LayoutCachedTop =1680
                    LayoutCachedWidth =18330
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =21480
                    Top =1680
                    Width =915
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label99"
                    Caption ="Order Qty"
                    LayoutCachedLeft =21480
                    LayoutCachedTop =1680
                    LayoutCachedWidth =22395
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =22440
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label100"
                    Caption ="FG UOM"
                    LayoutCachedLeft =22440
                    LayoutCachedTop =1680
                    LayoutCachedWidth =23700
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =20160
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label101"
                    Caption ="Trans Date"
                    LayoutCachedLeft =20160
                    LayoutCachedTop =1680
                    LayoutCachedWidth =21420
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11100
                    Top =840
                    Width =1320
                    TabIndex =5
                    ForeColor =0
                    Name ="btnWhereUsed"
                    Caption ="Where Used Run"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11100
                    LayoutCachedTop =840
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =1200
                    BorderWidth =1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =13800
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label111"
                    Caption ="CountPerUOM"
                    LayoutCachedLeft =13800
                    LayoutCachedTop =1680
                    LayoutCachedWidth =15060
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =780
                    Width =1320
                    TabIndex =2
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =780
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12540
                    Top =840
                    Width =1320
                    TabIndex =6
                    ForeColor =0
                    Name ="btnGetInventory"
                    Caption ="Get Inventory"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12540
                    LayoutCachedTop =840
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1200
                    BorderWidth =1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =23760
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label114"
                    Caption ="Revenue"
                    LayoutCachedLeft =23760
                    LayoutCachedTop =1680
                    LayoutCachedWidth =25020
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =4440
                    Top =1680
                    Width =3240
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label116"
                    Caption ="Oracle Description"
                    LayoutCachedLeft =4440
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =23940
                    Top =120
                    Width =720
                    Height =720
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000c40e0000c40e0000 ,
                        0x0000000000000000ffffffffffffffffffe1e1fe7777fe2c2cfe1414ff1414ff ,
                        0x1414ff1414ff1414ff1414ff1414ff1414ff1414ff1414ff1414ff1414ff1414 ,
                        0xff1414ff1414ff1414ff1414ff1414ff1414ff1414ff1414ff1414ff1414ff14 ,
                        0x14ff2c2cfe7777fee1e1feffffffffffffffffffffffffffffff9f9ffe0808fe ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0909fe9e9efeffffffffffff ,
                        0xffffffa0a0fe0101fe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0202fea1a1ffffffffe5e5ff0a0afe0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0b0bfee5e5ff8585ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff7e7efe3838fe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff3737fe2929fe0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff2929fe2929fe00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff2929fe2929fe0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff2020fe1c1cfe0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0101ff3f3fff0303fe0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0303fe4141fe0101ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff1717ff ,
                        0x1818ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0404 ,
                        0xfe9b9bfeffffffb5b5fe0909fe0000ff0000ff0000ff0000ff0000ff0000ff0a ,
                        0x0afeb4b4feffffff9e9efe0303fe0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff1515ff1616ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0101ff9a9afeffffffffffffffffffc0c0fe0a0afe0000ff00 ,
                        0x00ff0000ff0000ff0b0bfec2c2feffffffffffffffffff9a9afe0101ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff1414ff1515ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff2d2dfeffffffffffffffffffff ,
                        0xffffffffffc3c3fe0a0afe0000ff0000ff0a0afec2c2feffffffffffffffffff ,
                        0xffffffffffff2e2efe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff1414ff1414ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff01 ,
                        0x01ff9b9bfeffffffffffffffffffffffffffffffc1c1fe0808fe0a0afec2c2fe ,
                        0xffffffffffffffffffffffffffffff9d9dfe0101ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff1414ff1414ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0303fea5a5feffffffffffffffffffffffff ,
                        0xffffffc1c1fec2c2feffffffffffffffffffffffffffffffa5a5fe0303fe0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff1414ff1414ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0303fe ,
                        0xa4a4feffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffa5a5fe0303fe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff1414ff1414ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0303fea5a5feffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffa5a5fe0202ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff1414ff1414ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0303fea7a7 ,
                        0xfeffffffffffffffffffffffffffffffffffffa7a7ff0202fe0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff1414ff ,
                        0x1414ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0a0afec2c2feffffffffffffffffffffffffffffffffffffc0 ,
                        0xc0fe0808fe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff1414ff1414ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0b0bfec2c2feffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffc1c1fe0a0afe0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff1414ff1414ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0a0afec2c2feff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffc3c3fe ,
                        0x0a0afe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff1414ff1414ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0a0afec2c2feffffffffffffffffffffffffffffffa6a6fea6a6ffffffff ,
                        0xffffffffffffffffffffffffc0c0fe0909fe0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff1414ff1414ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0404feb3b3feffffffffffffffffffffffffffffff ,
                        0xa5a5fe0303fe0303fea5a5feffffffffffffffffffffffffffffffb4b4fe0202 ,
                        0xfe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff1414ff1616ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff2a2afefdfdffffffff ,
                        0xffffffffffffffffffa5a5fe0303fe0000ff0000ff0303fea4a4feffffffffff ,
                        0xfffffffffffffffefeff2828fe0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff1414ff1616ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff7d7dfeffffffffffffffffffa5a5fe0202ff0000ff0000ff0000 ,
                        0xff0000ff0303fea5a5feffffffffffffffffff7e7efe0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff1414ff1818ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0101ff7c7cfefdfdff9b9bfe0202 ,
                        0xfe0000ff0000ff0000ff0000ff0000ff0000ff0303fe9a9afefdfdff7c7cfe01 ,
                        0x01ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff1616ff ,
                        0x1f1ffe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff1d1dfe0101ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0101ff1e1efe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff1717ff2929fe0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff2222fe2929fe0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff2929fe2929fe0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff2929fe3f3ffe0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff3f3ffe8d8dff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff8d8dffefefff1515fe0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff1212fef2f2ffffffffbcbcfe0404fe0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0404feb6b6feffffff ,
                        0xffffffffffffbcbcfe1919fe0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x1717febdbdfefffffffffffffffffffffffffffffff2f2fe9a9afe4e4efe3434 ,
                        0xfe2e2efe2b2bfe2929fe2929fe2929fe2929fe2929fe2929fe2929fe2929fe29 ,
                        0x29fe2929fe2929fe2929fe2929fe2929fe2929fe2929fe2929fe2929fe2929fe ,
                        0x2a2afe3434fe5050fe9797fef0f0ffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="CloseIcon.bmp"

                    LayoutCachedLeft =23940
                    LayoutCachedTop =120
                    LayoutCachedWidth =24660
                    LayoutCachedHeight =840
                    BorderWidth =1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =18420
                    Top =1680
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label120"
                    Caption ="Line #"
                    LayoutCachedLeft =18420
                    LayoutCachedTop =1680
                    LayoutCachedWidth =20100
                    LayoutCachedHeight =1980
                End
            End
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="txOracleItem"
                    ControlSource ="Material"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =780
                    LayoutCachedTop =60
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7740
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txQuantity"
                    ControlSource ="MaterialQuantity_PerPrimaryUOM"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =60
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10320
                    Top =60
                    Width =810
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="txUsagePerQty"
                    ControlSource ="PrimaryUOM"

                    LayoutCachedLeft =10320
                    LayoutCachedTop =60
                    LayoutCachedWidth =11130
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11160
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="txUom"
                    ControlSource ="PrimaryUOM"

                    LayoutCachedLeft =11160
                    LayoutCachedTop =60
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9000
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="txBuyItem"
                    ControlSource ="MaterialQuanitty_PerOrder"

                    LayoutCachedLeft =9000
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12480
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="txComponentQty_M"
                    ControlSource ="Item_Status"

                    LayoutCachedLeft =12480
                    LayoutCachedTop =60
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15120
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="txStartUpScrap"
                    ControlSource ="Customer"

                    LayoutCachedLeft =15120
                    LayoutCachedTop =60
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2460
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txFinishedItem"
                    ControlSource ="AssemblyItemNumber"

                    LayoutCachedLeft =2460
                    LayoutCachedTop =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17100
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="Text104"
                    ControlSource ="OrderNum"

                    LayoutCachedLeft =17100
                    LayoutCachedTop =60
                    LayoutCachedWidth =18300
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21480
                    Top =60
                    Width =915
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =12
                    Name ="Text105"
                    ControlSource ="Quantity"

                    LayoutCachedLeft =21480
                    LayoutCachedTop =60
                    LayoutCachedWidth =22395
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22440
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="Text106"
                    ControlSource ="QTY_UOM"

                    LayoutCachedLeft =22440
                    LayoutCachedTop =60
                    LayoutCachedWidth =23700
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="Text107"
                    ControlSource ="TransactionDate"

                    LayoutCachedLeft =20160
                    LayoutCachedTop =60
                    LayoutCachedWidth =21420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13800
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="txCountPerUOM"
                    ControlSource ="CountPerUOM"

                    LayoutCachedLeft =13800
                    LayoutCachedTop =60
                    LayoutCachedWidth =15060
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =23760
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =14
                    Name ="Text115"
                    ControlSource ="Revenue"

                    LayoutCachedLeft =23760
                    LayoutCachedTop =60
                    LayoutCachedWidth =25020
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4440
                    Top =60
                    Width =3240
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txDescription"
                    ControlSource ="Description"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =60
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18420
                    Top =60
                    Width =1650
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =15
                    Name ="Text121"
                    ControlSource ="SoLineNumber"

                    LayoutCachedLeft =18420
                    LayoutCachedTop =60
                    LayoutCachedWidth =20070
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =660
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =300
                    Width =4380
                    Height =240
                    BorderColor =-2147483633
                    ForeColor =4144959
                    Name ="Label118"
                    Caption ="| Enter to Run  | Esc to cancel"
                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =540
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
Dim Materials As New clsMaterials
Dim Timer As New StopWatch

Private Sub Base_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

Private Sub btnClose_Click()
DoCmd.Close acForm, Me.name
End Sub




Private Sub btnExcelExport_Click()

On Error GoTo ErrorHandler
Dim fileName As String
fileName = "WhereUsedMaterials" & Format(Now, "yyyyMMddhhmmss")

'DoCmd.OutputTo acOutputTable, "tblGetBuyMaterials", acFormatXLS, Environ("USERPROFILE") & "\Desktop\BuyMaterials" & Format(Now, "yyyyMMddhhmmss") & ".xls", True

Export2XLS "tblGetWhereUsedMaterial", fileName

Beep
'MsgBox "File has successfully been exported to Desktop"

ErrorExit:
Exit Sub

ErrorHandler:
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit

End Sub

Private Sub btnGetInventory_Click()
DoCmd.OpenQuery "qryGetWhereUsedMaterialOnHnadInventory"
End Sub

Private Sub btnOpenInputs_Click()
DoCmd.OpenForm "frmMaterialWhereUsed", acFormDS, , , acFormEdit, acWindowNormal

End Sub

Private Sub btnRemoveFilter_Click()
Me.Filter = ""
Me.FilterOn = False
End Sub

Private Sub btnRun_Click()


DoCmd.RunSQL "DELETE FROM tblGetWhereUsedMaterial"

If CurrentProject.AllForms("frmMaterialWhereUsed").IsLoaded Then
    DoCmd.Close acForm, "frmMaterialWhereUsed", acSaveYes
End If

Timer.StartTimer

Materials.LoadData (2)

tb_FunctionTimer = Timer.EndTimer

Me.btnExcelExport.SetFocus
Me.Refresh
Me.Requery
Me.Repaint
Beep

End Sub



Private Sub btnWhereUsed_Click()

DoCmd.RunSQL "DELETE FROM tblGetWhereUsedMaterial"

If CurrentProject.AllForms("frmMaterialWhereUsed").IsLoaded Then
    DoCmd.Close acForm, "frmMaterialWhereUsed", acSaveYes
End If

Timer.StartTimer

DoCmd.OpenQuery "qryGetWhereUsedMaterial_Level1"

tb_FunctionTimer = Timer.EndTimer

Me.btnExcelExport.SetFocus
Me.Refresh
Me.Requery
Me.Repaint
Beep



End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

Private Sub Form_Load()

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE FROM tblGetWhereUsedMaterial"
Me.Refresh
Me.Requery

End Sub

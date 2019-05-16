Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =22
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =24240
    DatasheetFontHeight =10
    ItemSuffix =278
    Left =2445
    Top =780
    Right =30675
    Bottom =14880
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x229d1ec22bd9e440
    End
    RecordSource ="tblCableConstructionReferences_Local"
    Caption ="Create New Construction"
    OnDelete ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =900
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =240
                    Top =180
                    Width =6120
                    Height =615
                    FontSize =24
                    FontWeight =700
                    Name ="Label241"
                    Caption ="Create Cable Construction"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =240
                    LayoutCachedTop =180
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =795
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8340
                    Top =360
                    Width =2700
                    Height =315
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbBase_Copy"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblCableConstructionReferences.Base, tblCableConstructions.CablePasses FR"
                        "OM tblCableConstructionReferences INNER JOIN tblCableConstructions ON tblCableCo"
                        "nstructionReferences.BaseID = tblCableConstructions.BaseID WHERE (((tblCableCons"
                        "tructions.CablePasses)=1 Or (tblCableConstructions.CablePasses)=2)) ORDER BY tbl"
                        "CableConstructionReferences.Base;"
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =360
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =675
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6960
                            Top =360
                            Width =1020
                            Height =315
                            Name ="Label253"
                            Caption ="Copy Base"
                            LayoutCachedLeft =6960
                            LayoutCachedTop =360
                            LayoutCachedWidth =7980
                            LayoutCachedHeight =675
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12060
                    Top =240
                    Width =600
                    Height =600
                    TabIndex =1
                    Name ="btnCopy"
                    Caption ="Copy"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000075a0100075a0100 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffab ,
                        0xe8fe75dafe75dbfe75dbfe75dbfe75dbfe75dbfe75dbfe75dbfe75dbfe75dbfe ,
                        0x75dbfe75dbfe75dbfe75dbfe75dbfe75dbfe88dffeeefaffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffec8effe32c4fe8fdfff90dfff90dfff90dfff90dfff90dfff ,
                        0x90dfff90dfff90dfff90dfff90dfff90dfff90dfff90dfff90dfff90dffe59cf ,
                        0xfe76d7ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffade4fe5dcbffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffade4fe5dcbffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xaee2fe5fc7ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffaee2fe5fc7ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffafe0fe61c2ffffffffffffffffffffffffffffff ,
                        0xfff1f9ffe4f4ffe4f4ffe4f4ffe4f4ffe4f4ffe4f4ffe4f4ffe4f4ffe4f4ffe4 ,
                        0xf4fe9ed9fe5ac0ffe4f4ffe4f4ffe4f4fee5f4fffbfdffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffb0ddfe62beffffff ,
                        0xffffffffffffffffffffbae2fe33aaff2fa9ff2fa9ff2fa9ff2fa9ff2fa9ff2f ,
                        0xa9ff2fa9ff2fa9ff2fa9ff2fa9ff2fa9ff2fa9ff2fa9ff2fa9ff2fa9ff2fa9fe ,
                        0x4cb4ffeff7ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffb1dbfe64b9fffffffffffffffffffffffffe68bafea3d5ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffdfffe58b4feb5ddffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffb2d9fe66b5ffffffffffffffffffffff ,
                        0xfffe66b5ffb2d9ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffe66b5ffb2d9ffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffb3d7fe68 ,
                        0xb0ffffffffffffffd9eaffd9eafe5fabffb3d7ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffe68b0ffb3d7ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffb4d4fe6aacffffffffffffff3a91fe4396fe3c92ffb4d4ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffe6aacffb4d4ffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffb5d2fe6ca8ffffffffffffff ,
                        0x3c8bffe6effe6ca8ffb5d2ffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffe6ca8ffb5 ,
                        0xd2ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xb6d0fe6ea3ffffffffffffff3f85ffe6effe6ea3ffb6d0ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffe6ea3ffb6d0ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffb6cefe709fffffffffffffff417ffe6093fe4984 ,
                        0xffb6cefffffffffffffef3f6febfd4febfd4febfd4febfd4febfd4febfd4febf ,
                        0xd4febfd4febfd4feccdcfffffffffffffffffffe709fffb6ceffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffb7ccfe729affffff ,
                        0xffffffffc3d5ffc3d5fe6390ffb7ccffffffffffffffeceaffb7b0ffb7b0ffb7 ,
                        0xb0ffb7b0ffb7b0ffb7b0ffb7b0ffb7b0ffb7b0febfb9fffffffffffffffffffe ,
                        0x729affb7ccffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffb8c9fe7496fffffffffffffe97affe97affe5a82ffb8c9ffffffffff ,
                        0xffffeee3ffbe93ffbe93ffbe93ffbe93ffbe93ffbe93ffbe93ffbe93ffbe93fe ,
                        0xc59ffffffffffffffffffffe7496ffb8c9ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffb9c7fe7691fffffffffffffe98adfe98 ,
                        0xadfe5c7dffb9c7fffffffffffffef9f3feddb5feddb5feddb5feddb5feddb5fe ,
                        0xddb5feddb5feddb5feddb5fee5c7fffffffffffffffffffe7691ffb9c7ffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffbac5fe78 ,
                        0x8dffffffffffffffc6cfffc6cffe6a81ffbac5fffffffffffffff0f3ffc6cfff ,
                        0xc6cfffc6cfffc6cfffc6cfffc6cfffc6cfffc6cfffc6cffecdd4ffffffffffff ,
                        0xfffffffe788dffbac5ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffbbc3fe7a88fffffffffffffe6e7ffe6e7ffe5568ffbbc3ff ,
                        0xfffffffffffedadefe6e7ffe6e7ffe6e7ffe6e7ffe6e7ffe6e7ffe6e7ffe6e7f ,
                        0xfe6e7ffe808ffffffffffffffffffffe7a88ffbbc3ffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffbcc0fe7c84ffffffffffffff ,
                        0xf3f4fff3f4fe7981ffbcc0fffffffffffffffbfbfff3f4fff3f4fff3f4fff3f4 ,
                        0xfff3f4fff3f4fff3f4fff3f4fff3f4fef4f5fffffffffffffffffffe7c84ffbc ,
                        0xc0ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xbdbffe7d7fffffffffffffff5356ff5356ff5356ffbdbfffffffffffffffd3d3 ,
                        0xff5356ff5356ff5356ff5356ff5356ff5356ff5356ff5356ff5356fe686affff ,
                        0xfffffffffffffffe7d7fffbdbfffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffbfbcfe7f7bfffffffffffffff3f3fff3f3fff3f3 ,
                        0xfffbfbfffffffffffffffbfbfff3f3fff3f3fff3f3fff3f3fff3f3fff3f3fff3 ,
                        0xf3fff3f3fff3f3fef5f5fffffffffffffffffffe7f7bffbfbcffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffbfbafe8176ffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffedcd9fe776bfe776bfe77 ,
                        0x6bfe776bfe776bfe776bfe776bfe776bfe776bfe887dfffffffffffffffffffe ,
                        0x8176ffbfbaffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffc0b7fe8372ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffff1efffcac4ffcac4ffcac4ffcac4ffcac4ffcac4ffcac4ffcac4ffcac4fe ,
                        0xd1cbfffffffffffffffffffe8372ffc0b7ffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffec4b9fe7c62fffdfdffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffe856dffc1b5ffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffff1effe76 ,
                        0x53ff6038ff6038ff6038ff6038ff6038ffc2b3ffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffea48dfe623b ,
                        0xff6038ff6038feccbfffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffcfbfeebe5ffeae5ffeae5ffeae5fe845effc3b1ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffefebff6332fee0d7fe9371fe835dfffcfbffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffe8b60ffc4afffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffebe4ff652cfec9b5fe6e39ffebe4ffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffe8d5cffc5acffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffebe3ff68 ,
                        0x26fe6f31ffcfbaffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffea173 ,
                        0xfe8d55ffb38dffb38dffb38dffb38dffb38dffb38dffb38dffb38dffb38dffb3 ,
                        0x8dffb38dffb38dfeab80ff6a20feab81ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffef4eefeae83fea370fea370fea370fea370fea370fea3 ,
                        0x70fea370fea370fea370fea370fea370fea370fea370feb892fffcfbffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Copy"
                    Picture ="Copy.bmp"

                    LayoutCachedLeft =12060
                    LayoutCachedTop =240
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =23160
                    Top =120
                    Width =780
                    Height =660
                    TabIndex =2
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

                    LayoutCachedLeft =23160
                    LayoutCachedTop =120
                    LayoutCachedWidth =23940
                    LayoutCachedHeight =780
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11100
                    Top =300
                    Width =426
                    Height =426
                    TabIndex =3
                    Name ="btnOpenBase"
                    Caption ="btnOpenSub"
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
                    ControlTipText ="Open Construction"
                    Picture ="OpenIcon.bmp"

                    LayoutCachedLeft =11100
                    LayoutCachedTop =300
                    LayoutCachedWidth =11526
                    LayoutCachedHeight =726
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =11220
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =5400
                    Top =4320
                    Width =705
                    Height =315
                    Name ="Label48"
                    Caption ="Base"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =4320
                    LayoutCachedWidth =6105
                    LayoutCachedHeight =4635
                End
                Begin Label
                    OverlapFlags =93
                    Left =5400
                    Top =2580
                    Width =960
                    Height =315
                    Name ="Label49"
                    Caption ="Design Code"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =2580
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =2895
                End
                Begin Label
                    OverlapFlags =93
                    Left =5400
                    Top =5220
                    Width =705
                    Height =315
                    Name ="Label52"
                    Caption ="TB Type"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =5220
                    LayoutCachedWidth =6105
                    LayoutCachedHeight =5535
                End
                Begin Label
                    OverlapFlags =93
                    Left =5400
                    Top =3010
                    Width =1260
                    Height =315
                    Name ="Label17"
                    Caption ="Design Code Ext"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =3010
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =3325
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =4320
                    Width =3660
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    ForeColor =-2147483640
                    Name ="txBase"
                    ControlSource ="Base"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =4320
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =4635
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =2580
                    Width =3660
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =-2147483640
                    Name ="txDesignCode"
                    ControlSource ="DesignCode"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =2580
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =2895
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =6900
                    Top =5190
                    Width =3660
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"40\""
                    Name ="cbTBType"
                    ControlSource ="TBType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblCableTightBufferReference.TBType, tblCableTightBufferReference.TBTypeI"
                        "D FROM tblCableTightBufferReference ORDER BY tblCableTightBufferReference.TBType"
                        ";"
                    ColumnWidths ="2880;1440;1440"
                    DefaultValue ="Null"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =5190
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =5505
                End
                Begin Label
                    OverlapFlags =93
                    Left =5400
                    Top =3865
                    Width =1380
                    Height =405
                    Name ="Label202"
                    Caption ="New Product Code"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =3865
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =4270
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =9170
                    Width =1500
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =25
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txCableFamily"
                    ControlSource ="CableFamily"
                    Format =">"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =9170
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =9485
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5400
                            Top =9170
                            Width =960
                            Height =315
                            Name ="Label204"
                            Caption ="Cable Family"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =9170
                            LayoutCachedWidth =6360
                            LayoutCachedHeight =9485
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =9594
                    Width =1500
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =27
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txCableLevel1"
                    ControlSource ="CableLevel1"
                    Format =">"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =9594
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =9909
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5400
                            Top =9594
                            Width =1080
                            Height =315
                            Name ="Label206"
                            Caption ="Cable Level 1"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =9594
                            LayoutCachedWidth =6480
                            LayoutCachedHeight =9909
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =3000
                    Width =3660
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txDesignCodeExt"
                    ControlSource ="DesignExtension"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =3000
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =3315
                End
                Begin Label
                    OverlapFlags =93
                    Left =5400
                    Top =1300
                    Width =1020
                    Height =315
                    Name ="Label211"
                    Caption ="Cable Passes"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =1300
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =1615
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5820
                    Top =7822
                    Width =1380
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =21
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txNumSubs_2Pass"
                    ControlSource ="NumSubPositions"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =7822
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =8137
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5820
                            Top =7470
                            Width =1380
                            Height =315
                            Name ="lblNumPosition_2Pass"
                            Caption ="# Total Positions"
                            LayoutCachedLeft =5820
                            LayoutCachedTop =7470
                            LayoutCachedWidth =7200
                            LayoutCachedHeight =7785
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7380
                    Top =7822
                    Width =1380
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =22
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txNumFillers_2Pass"
                    ControlSource ="NumSubFillers"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7380
                    LayoutCachedTop =7822
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =8137
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7380
                            Top =7470
                            Width =1380
                            Height =315
                            Name ="lblNumFillers_2Pass"
                            Caption ="# Fillers"
                            LayoutCachedLeft =7380
                            LayoutCachedTop =7470
                            LayoutCachedWidth =8760
                            LayoutCachedHeight =7785
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8940
                    Top =7822
                    Width =1380
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =23
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txNumCu_2Pass"
                    ControlSource ="NumCopperUnits"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8940
                    LayoutCachedTop =7822
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =8137
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8940
                            Top =7470
                            Width =1380
                            Height =315
                            Name ="lblCuUnits_2Pass"
                            Caption ="# Copper Units"
                            LayoutCachedLeft =8940
                            LayoutCachedTop =7470
                            LayoutCachedWidth =10320
                            LayoutCachedHeight =7785
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =1300
                    Width =1140
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cbCablePasses"
                    RowSourceType ="Value List"
                    RowSource ="1;2"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =1300
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =1615
                End
                Begin Label
                    OverlapFlags =93
                    Left =5400
                    Top =3430
                    Width =1080
                    Height =315
                    Name ="Label226"
                    Caption ="Design Type"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =3430
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =3745
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =7200
                    Left =6900
                    Top =3420
                    Width =3660
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"140\""
                    Name ="cbDesignType"
                    ControlSource ="DesignTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblDesignTypes.DesignTypeID, tblDesignTypes.SalesForceDescription, tblDes"
                        "ignTypes.Product_Pricing_Group, tblDesignTypes.Active, tblDesignTypes.DesignType"
                        "GroupID, tblDesignTypes.DesignTypeDesc FROM tblDesignTypes ORDER BY tblDesignTyp"
                        "es.SalesForceDescription;"
                    ColumnWidths ="0;4320;2880"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =6900
                    LayoutCachedTop =3420
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =3735
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =93
                    Left =5400
                    Top =780
                    Width =2880
                    Height =405
                    FontSize =14
                    FontWeight =600
                    Name ="Label229"
                    Caption ="Cable Type"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =780
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =1185
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =12960
                    Left =6900
                    Top =3850
                    Width =840
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cbFamilyID"
                    ControlSource ="FamilyID"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateConstructionFamilyID"
                    ColumnWidths ="1440;4680;2520;2160;2160"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =3850
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =4195
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =6890
                    Width =1080
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cbSubTemplate_2Pass"
                    ControlSource ="UnitIDTypeNumber"
                    RowSourceType ="Value List"
                    RowSource ="1;2"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =6890
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =7205
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5400
                            Top =6890
                            Width =1260
                            Height =315
                            Name ="lblSubTemplate"
                            Caption ="Sub Template "
                            LayoutCachedLeft =5400
                            LayoutCachedTop =6890
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =7205
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    Left =5400
                    Top =1730
                    Width =1335
                    Height =315
                    Name ="Label232"
                    Caption ="Test Construction"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =1730
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =2045
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6900
                    Top =1730
                    Width =1140
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cbTest"
                    ControlSource ="ReleasedDesign"
                    RowSourceType ="Value List"
                    RowSource ="0;\"Yes\";1;\"No\""
                    ColumnWidths ="0;1440"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =1730
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =2045
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =5760
                    Left =6900
                    Top =6030
                    Width =3660
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbSubunit_2Pass"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateConstructionSubunit"
                    ColumnWidths ="2520;1080;1080;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =6030
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =6345
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5400
                            Top =6030
                            Width =1260
                            Height =315
                            Name ="lblSubunit"
                            Caption ="Subunit "
                            LayoutCachedLeft =5400
                            LayoutCachedTop =6030
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =6345
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =10740
                    Top =3649
                    Width =960
                    Height =645
                    TabIndex =7
                    Name ="btnAddNPR"
                    Caption ="Add New NPR"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000130b0000130b0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xc4c4c48989898080808080808080808080808080808080808080808080808080 ,
                        0x8080808080808080808080808080808080808080808080808080808080808080 ,
                        0x8080808080808080808080808080898989c5c5c5ffffffffffffffffffffffff ,
                        0xffffffffffffefefef4747470000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x474747efefeffffffffffffffffffffefefe4646460000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000474747ffffffffffffffffffc4c4c40000 ,
                        0x000000000202026161617f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f ,
                        0x7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f ,
                        0x7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f616161020202000000000000c5c5 ,
                        0xc5ffffffffffff898989000000000000616161ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xff616161000000000000898989ffffffffffff7f7f7f0000000000007f7f7fff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f000000000000808080ffffffffffff7f ,
                        0x7f7f0000000000007f7f7fffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f00000000 ,
                        0x0000808080ffffffffffff7f7f7f0000000000007f7f7fffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffff7f7f7f000000000000808080ffffffffffff7f7f7f000000000000 ,
                        0x7f7f7fffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffff7f7f7f000000000000808080ffffff ,
                        0xffffff7f7f7f0000000000007f7f7fffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x000000000000808080ffffffffffff7f7f7f0000000000007f7f7fffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff7f7f7f000000000000808080ffffffffffff7f7f7f0000 ,
                        0x000000007f7f7fffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff7f7f7f0000000000008080 ,
                        0x80ffffffffffff7f7f7f0000000000007f7f7fffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xff7f7f7f000000000000808080ffffffffffff7f7f7f0000000000007f7f7fff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f000000000000808080ffffffffffff7f ,
                        0x7f7f0000000000007f7f7fffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f00000000 ,
                        0x0000808080ffffffffffff7f7f7f0000000000007f7f7fffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffd2d2d2b9b9 ,
                        0xb9fefefeffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffff838383000000000000838383ffffffffffff7f7f7f000000000000 ,
                        0x7f7f7fffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffd2d2d2040404000000636363fdfdfdffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffd8d8d81e1e1e1f1f1fd9d9d9ffffff ,
                        0xffffff7f7f7f0000000000007f7f7fffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffb8b8b8000000000000000000636363fd ,
                        0xfdfdffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffff7f7f7f0000000000007f7f7fffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffefefe62 ,
                        0x6262000000000000000000636363fdfdfdffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f0000 ,
                        0x000000007f7f7fffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffdfdfd626262000000000000000000636363fdfdfd ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffff7f7f7f0000000000007f7f7fffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffdfdfd626262 ,
                        0x000000000000000000636363fdfdfdffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffff7f7f7f0000000000007f7f7fff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffdfdfd626262000000000000000000636363fdfdfdffff ,
                        0xfffffffffffffffffffffffffffdfdfd9e9e9e9e9e9efdfdfdffffffffffff7f ,
                        0x7f7f0000000000007f7f7fffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffdfdfd6262620000 ,
                        0x00000000000000636363fdfdfdffffffffffffffffffffffff9e9e9e00000000 ,
                        0x00009f9f9fffffffffffff7f7f7f0000000000007f7f7fffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffdfdfd626262000000000000000000636363fdfdfdffffffff ,
                        0xffffffffff7f7f7f000000000000808080ffffffffffff7f7f7f000000000000 ,
                        0x7f7f7fffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffdfdfd62626200000000 ,
                        0x0000000000636363fdfdfdffffffffffff7f7f7f000000000000808080ffffff ,
                        0xffffff7f7f7f0000000000007f7f7fffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffdfdfd626262000000000000000000636363fdfdfdffffff7f7f7f ,
                        0x000000000000808080ffffffffffff7f7f7f0000000000007f7f7fffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffdfdfd626262000000000000 ,
                        0x000000636363fdfdfd7f7f7f000000000000808080ffffffffffff7f7f7f0000 ,
                        0x000000007f7f7fffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffdfdfd6262620000000000000000006363637e7e7e0000000000008080 ,
                        0x80ffffffffffff7f7f7f0000000000007f7f7fffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffdfdfd6262620000000000000000 ,
                        0x00131313000000000000808080ffffffffffff898989000000000000616161ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffdfdfd626262000000000000000000000000000000808080ffffffffffffc4 ,
                        0xc4c40000000000000202026161617f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f ,
                        0x7f7f7f7f7f7f7f7f7f7f7f7f838383d9d9d9fffffffffffffffffffffffffdfd ,
                        0xfd9e9e9e7f7f7f7f7f7f7f7f7f7f7f7f7d7d7d12121200000000000000000000 ,
                        0x0000808080fffffffffffffefefe464646000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000001f1f1fffff ,
                        0xffffffffffffffffffff9e9e9e00000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000808080ffffffffffffffffffeeeeee464646 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000001f1f1fffffffffffffffffffffffff9e9e9e00000000000000 ,
                        0x00000000000000000000000000000000000000000000000000009f9f9fffffff ,
                        0xfffffffffffffffffffefefec4c4c48989897f7f7f7f7f7f7f7f7f7f7f7f7f7f ,
                        0x7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f838383d8d8d8ffffffffffffffffffff ,
                        0xfffffdfdfd9e9e9e7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f ,
                        0x7f7f7f9f9f9ffdfdfdffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Create New NPC"
                    Picture ="BigOpen.bmp"

                    LayoutCachedLeft =10740
                    LayoutCachedTop =3649
                    LayoutCachedWidth =11700
                    LayoutCachedHeight =4294
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =10740
                    Top =4414
                    Width =960
                    Height =480
                    TabIndex =10
                    Name ="btnBurnListing"
                    Caption ="Add Burn Listing"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000c40e0000c40e0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffbfbfbdedee1b6b5c99090ba87 ,
                        0x87bb8686bc8c8ebaadadc6d7d7d8f8f8f8ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffe8e7e88a8bb536 ,
                        0x35990b0ca00001a90000ae0000ae0101b10100b30001b3090aac302fa47e7fb7 ,
                        0xdbdddfffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff1 ,
                        0xf2f38688b119199a0101a60101aa0101ad0001ae0001b10101b30100b50101b6 ,
                        0x0101b80100b90100ba0102b81211ad7576b6eaececffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffd7d8dd3a3a9a0202a10101a80100ab0101ab0101ad0000af ,
                        0x0000b10101b30100b40000b60001b80000b90100bb0100bd0101bf0101bc292a ,
                        0xaac3c4d3ffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffc4c6d51819950201a30201a80100a8 ,
                        0x0000a90101ac0101ad0000ae0605b12626bd2929bf0a08b60001b80100b90100 ,
                        0xbc0001bd0001be0000bf0101bf0f11b3abaccfffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffccccd8161596 ,
                        0x0001a20100a50100a70100a80000aa0304aa5957c5bdbbe6f6f6fbffffffffff ,
                        0xfffbfbfdc7c6ed6968d30606b70001bd0001be0000bf0100c10100c30c0cb6ba ,
                        0xb9d5ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffe8e8ec25269600009f0000a30100a50100a70101a81313aeb2b2e3ffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffd0cff02423c30101bf00 ,
                        0x00bf0100c10100c30100c31617b3dadce4ffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffefefe6162a701019c0000a00000a30100a50101 ,
                        0xa50b0babcacaeaffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffe6e5f72021c30201c00100c10100c30200c50201c64847b8f9fbf9 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffc6c7d70706960101 ,
                        0xa00000a00000a30100a50202a59290d3ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffc2c1ed0403bf0100c10100c3 ,
                        0x0200c60101c60304c2a8a8d0ffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5152a101009d0101a00000a00000a30100a51e1cacf9f9fcffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffff4a49d00100c10100c30200c60100c80102c93433b3fcfcfcffffffffff ,
                        0xffffffffffffffffffffffffffe5e5ea0e0d9200009e0001a00000a00000a301 ,
                        0x00a56a69c6ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffff9c9be40302c20100c30200c60100c80101 ,
                        0xc90503c3c9cadfffffffffffffffffffffffffffffffffffff9d9dc301019900 ,
                        0x009e0001a00000a00000a30100a59e9cdbffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffceccf00102 ,
                        0xc10100c30200c60100c80101c80101c87d7fcbffffffffffffffffffffffffff ,
                        0xffffffffff6363af01009b00009d0000a00000a00000a30100a5bebde7ffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffff3f3fb8484da7f7e ,
                        0xdaeeeef9ffffffd9d8f40201c10100c30200c60101c80101c90101ca3f40bdff ,
                        0xffffffffffffffffffffffffffffffffff35369c01019c00009d0001a00000a0 ,
                        0x0000a30100a5a3a3ddffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff6665d10202ba0202bb3735c8fbfafdb5b4ea0102c20100c30200c601 ,
                        0x00c80101c90101cb1d1bbcfbfcfcffffffffffffffffffffffffffffff1f2195 ,
                        0x01009c00009d0001a00000a00000a30100a57573caffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffff2c2ac30000b90100bc0302beb6b5e960 ,
                        0x5ed50100c10100c30200c60100c80101c90101cb0f0ebff3f4f6ffffffffffff ,
                        0xffffffffffffffffff1b1b9201009c00009d0001a00000a00000a30100a52725 ,
                        0xb0fdfdfeffffffe1e0f2acabe1fcfcfeffffffffffffffffffffffff3a39c500 ,
                        0x00b90100bc0001bd5754d10505c00100c10100c30200c60100c70101c90101cb ,
                        0x0a0bbef2f2f7ffffffffffffffffffffffffffffff1e1e9501009c00009d0000 ,
                        0xa00000a00000a30100a50202a69e9cd9ffffff8584d50102ac6a6acdffffffff ,
                        0xffffffffffffffff6968d20100ba0100bc0001bd0102be0201bf0100c10100c3 ,
                        0x0200c60100c80101c90101cb0e0dbff3f4f6ffffffffffffffffffffffffffff ,
                        0xff3938a101009c00009d0000a00000a00000a30100a50101a61514ace2e2f290 ,
                        0x90d90101ad0202aeb0b0e4ffffffffffffffffff9190de0101b80100bc0001bd ,
                        0x0001be0000bf0100c10100c30200c60100c80101c90101cb1c19bffcfcfcffff ,
                        0xffffffffffffffffffffffffff6361b201019c00009e0001a00000a00000a301 ,
                        0x00a50100a70202a83333b9b0b0e00101ad0101ae3e3dc1ffffffffffffffffff ,
                        0xb1aee60101b90100bc0001bd0001be0000bf0100c10100c30200c60100c80101 ,
                        0xc90101cb4140c6ffffffffffffffffffffffffffffffffffffa0a1cd01009a00 ,
                        0x009d0001a00000a00000a30100a50100a70100a80101aa3131b61213b00001ae ,
                        0x0404aff2f1faffffffffffffb5b3e60101b90100bc0001bd0001be0000bf0100 ,
                        0xc10100c30200c60100c80000c80000c77b7ad1ffffffffffffffffffffffffff ,
                        0xffffffffffe8e9f10a0a9600009d0001a00000a00000a30100a50100a70100a8 ,
                        0x0000a90101ab0101ac0000af0101b0c0bfeaffffffffffffa4a3e20202b90100 ,
                        0xbc0001bd0001be0000bf0100c10100c30200c60100c80101c80402c7cacae6ff ,
                        0xffffffffffffffffffffffffffffffffffffffff5758ae00019f0000a00000a0 ,
                        0x0100a30100a50200a70100a80000a90101ab0101ad0000af0101b0bbb9e9ffff ,
                        0xffffffff6666d30200ba0100bc0001bd0001be0000bf0100c10100c30200c601 ,
                        0x01c70202ca3234c3fbfbfcffffffffffffffffffffffffffffffffffffffffff ,
                        0xc7c7de04049b00009f0000a00000a30100a50100a70100a80000a90101ab0101 ,
                        0xad0001af0202b0e2e2f4ffffffdfdff50909b90000ba0100bc0001bd0001be00 ,
                        0x00bf0100c10100c30200c60101c70302c7aaabdcffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffeff6463b40101a20000a00000a30100a50100 ,
                        0xa70100a80000a90101ab0101ad0101ae3231bdfffffff3f3fb3535c10102b800 ,
                        0x00b90100bc0001bd0001be0000bf0100c10100c30200c60201c94342c5fdfdfd ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffededf32221 ,
                        0xa00202a40100a30100a50100a70100a80000a90101ab0101ad0202aeababe1d3 ,
                        0xd3f02f2fbf0101b50001b80000b90100bc0001bd0001be0000bf0100c10100c3 ,
                        0x0201c81313c2dbdceeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffd3d3e51312a10202a50000a40100a70100a80000a901 ,
                        0x01ab0403ad6160c77575d30707b30200b50001b60001b80000ba0100bc0001bd ,
                        0x0001be0000bf0101c10100c60707c2bbbbe3ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffcfcee51515a401 ,
                        0x01aa0100a80100a80000a90101ab0404ad0808b00202b00101b30100b40000b6 ,
                        0x0001b80000ba0100bc0001bd0101be0100c10201c5090ac0b4b4e3ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffdcdcec3737ad0101ab0001ab0000aa0101aa0101ad0000af ,
                        0x0001b10101b30100b40000b60001b80000b90100bc0001bd0101c20201c62725 ,
                        0xc0cbcbe9ffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffff7f7fa8b8bc81414a8 ,
                        0x0102b20000b10001af0000af0001b10101b30100b40000b60101b90101bb0000 ,
                        0xc10101c40c0cbd7777cdf4f4faffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffeeeff79192cd3635b30504ae0203b50102b80001ba0100 ,
                        0xbb0101bb0202be0504b92a2bbc8582d3e5e5f2ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe6e6 ,
                        0xf1bbbae19797d5908fd68f8fd69494d8b1b1dce0e1edffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add Burn Listing"
                    Picture ="burnlisting.bmp"

                    LayoutCachedLeft =10740
                    LayoutCachedTop =4414
                    LayoutCachedWidth =11700
                    LayoutCachedHeight =4894
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =10440
                    Top =9590
                    Width =1320
                    TabIndex =26
                    Name ="btnSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000c40e0000c40e0000 ,
                        0x0000000000000000fffffffffffffffffffdfdfdfdfdfdffffffffffffffffff ,
                        0xfefefefdfdfdfefefefefffeeaf6e8b3dfac87cd7c6ac05c56b9464bb53b4bb5 ,
                        0x3b57b9476bc15d8ace7fb5e0aeebf7eafefffefefefefdfdfdfefefeffffffff ,
                        0xfffffffffffdfdfdfdfdfdffffffffffffffffffffffffffffffffffffffffff ,
                        0xf9f9f9fbfbfbfffffffffffffffffffbfbf9b9deb372c3644db63d4ab5394ab5 ,
                        0x394ab3394ab2394ab5394ab5394ab5394ab5394ab2394ab3394db53d73c566bf ,
                        0xe4b9fbfcfbf8f8f8fdfdfdfffffffffffffffffffbfbfbfbfbfbffffffffffff ,
                        0xfefefefefefefffffffffffffefefef6f6f6f8f8f8fefffecdeac866c0584ab3 ,
                        0x3949af3a4ab3394ab5394ab5394ab5394ab13a49b03a4ab5394ab5394ab5394a ,
                        0xb43949b03a4ab2394ab5394ab53967c059cbe7c7f5f5f5fdfdfdffffffffffff ,
                        0xfffffff7f7f7fbfbfbfffffffefefefbfbfbfefefefffffffffffffefefef0f3 ,
                        0xef8ccb824bb53a4ab5394ab5394ab33949af3a4ab3394ab5394ab5394ab53949 ,
                        0xb13a49b03a4ab5394ab5394ab5394ab43949af3a4ab2394ab5394ab5394cb53b ,
                        0x8fce85f0f3effdfdfdfffffffffffffffffffbfbfbfdfdfdfffffffbfbfbf6f6 ,
                        0xf6fefefeffffffedf7eb6dc35f49b03a49b13a4ab5394ab5394ab5394ab33949 ,
                        0xaf3a4ab3394ab5394ab5394ab53949b13a49b03a4ab5394ab5394ab5394ab439 ,
                        0x49af3a4ab2394ab5394ab5394ab5396fc162e7eee6fdfdfdffffffffffffffff ,
                        0xfffdfdfdfffffffffffff9f9f9f4f4f4ecf6eb65c0564ab5394ab43949af3a49 ,
                        0xb13a4ab5394ab5394ab5394ab33949af3a4ab3394ab5394ab5394ab53949b13a ,
                        0x49b03a4ab5394ab5394ab5394ab43949af3a4ab2394ab5394ab5394ab53966bd ,
                        0x58e7eee6fdfdfdfffffffffffffffffffffffffffffffffffff0f5f06abc5e4a ,
                        0xb4394ab5394ab5394ab53949b03a49b13a4ab5394ab5394ab5394ab33949af3a ,
                        0x4ab3394ab5394ab5394ab53949b13a49b03a4ab5394ab5394ab5394ab43949af ,
                        0x3a4ab2394ab5394ab5394ab5396fc162f0f3effdfdfdffffffffffffffffffff ,
                        0xfffffefefe8dd0834ab23949af3a4ab4394ab5394ab5394ab53949af3a49b13a ,
                        0x4ab5394ab5394ab5394ab33949af3a4ab3394ab5394ab5394ab53949b13a49b0 ,
                        0x3a4ab5394ab5394ab5394ab43949af3a4ab2394ab5394ab5394ab5398fce85f5 ,
                        0xf5f5fdfdfdfffffffdfdfdfdfdfdc9e8c44ab4394ab5394ab23949af3a4ab439 ,
                        0x4ab5394ab5394ab53949b03a54b545c1e6bbb3e0ac4eb63d4ab33949af3a4ab3 ,
                        0x394ab5394ab5394ab53949b13a49b03a4ab5394ab5394ab5394ab43949af3a4a ,
                        0xb2394ab5394ab5394cb53bcbe7c7f8f8f8fefefefefefef5f8f563bd554ab539 ,
                        0x4ab5394ab5394ab23949af3a4ab4394ab5394ab53954b944d0e6ccf9f9f9ffff ,
                        0xffc3e6bd4db63c4ab33949af3a4ab3394ab5394ab5394ab53949b13a49b03a4a ,
                        0xb5394ab5394ab5394ab43949af3a4ab2394ab5394ab53967c059fbfcfbfdfdfd ,
                        0xffffffbce2b64ab03a4ab3394ab5394ab5394ab5394ab23949af3a4ab4394fb7 ,
                        0x3fd4edd0fefefef6f6f6f9f9f9ffffffc0e5ba4cb63b4ab33949af3a4ab3394a ,
                        0xb5394ab5394ab53949b13a49b03a4ab5394ab5394ab5394ab43949af3a4ab239 ,
                        0x4ab5394ab539bfe4b9fefefefdfdfd6ac25c4ab43949af3a4ab3394ab5394ab5 ,
                        0x394ab5394ab2394fb140c9e8c4fffffffffffffefefef6f6f6f9f9f9ffffffb2 ,
                        0xe0ab4cb63b4ab33949af3a4ab3394ab5394ab5394ab53949b13a49b03a4ab539 ,
                        0x4ab5394ab5394ab43949af3a4ab2394ab53973c566fefffee2f2df4bb53a4ab5 ,
                        0x394ab43949af3a4ab3394ab5394ab5394eb63dc4e3bff6f6f6fefefeffffffff ,
                        0xfffffefefef6f6f6f9f9f9ffffffb3e0ac4bb53a4ab33949af3a4ab3394ab539 ,
                        0x4ab5394ab53949b13a49b03a4ab5394ab5394ab5394ab43949af3a4ab2394db5 ,
                        0x3debf7eab0dea94ab5394ab5394ab5394ab43949af3a4ab33950b840c9e9c4ff ,
                        0xfffff9f9f9f6f6f6fefefefffffffffffffefefef6f6f6f9f9f9ffffffb2dfab ,
                        0x4cb63b4ab33949af3a4ab3394ab5394ab5394ab53949b13a49b03a4ab5394ab5 ,
                        0x394ab5394ab43949b03a4ab339b5e0ae87cc7c4ab3394ab5394ab5394ab5394a ,
                        0xb4394db03ec6e5c1fffffffffffffffffff9f9f9f6f6f6f9fcf9fefefeffffff ,
                        0xfefefef6f6f6f9f9f9fefffeb2e0ab4bb53a4ab33949af3a4ab3394ab5394ab5 ,
                        0x394ab53949b13a49b03a4ab5394ab5394ab5394ab4394ab2398ace7f6ac15c4a ,
                        0xb3394ab1394ab5394ab53950b73fc7e7c2f5f5f5fbfbfbffffffffffffffffff ,
                        0xf1f6f177c26b8ed083fcfefcfffffffefefef6f6f6f9f9f9ffffffb2e0ab4cb6 ,
                        0x3b4ab33949af3a4ab3394ab5394ab5394ab53949b13a49b03a4ab5394ab5394a ,
                        0xb5394ab5396bc15d55b9454ab5394ab13a49b03a4ab53988ce7dfffffffdfdfd ,
                        0xf5f5f5fbfbfbfffffff7fcf778c86c4ab23949af3a8dd083fdfefdfffffffefe ,
                        0xfef6f6f6f9f9f9fefffeb2dfab4bb53a4ab33949af3a4ab3394ab5394ab5394a ,
                        0xb53949b13a49b03a4ab5394ab5394ab53957b9474bb53a4ab5394ab53949b03a ,
                        0x49b03a5cbc4ce6f5e4fffffffdfdfdf5f5f5f3f8f37ac96e4ab5394ab5394ab2 ,
                        0x3949af3a8dd082fdfefcfffffffefefef6f6f6f9f9f9ffffffb3e0ac4cb63b4a ,
                        0xb33949af3a4ab3394ab5394ab5394ab53949b13a49b03a4ab5394ab5394bb53b ,
                        0x4db63c4ab5394ab5394ab53949b03a49b03a5dbd4ee7f5e5fffffff5faf474c1 ,
                        0x694ab3394ab5394ab5394ab5394ab23949af3a8cd082fdfefcfffffffefefef6 ,
                        0xf6f6f9f9f9ffffffb2dfab4bb63a4ab33949af3a4ab3394ab5394ab5394ab539 ,
                        0x49b13a49b03a4ab5394eb63d57b9474ab5394ab5394ab5394ab53949b03a49b0 ,
                        0x3a60be51b5e1ae77c76a4ab43949af3a4ab3394ab5394ab5394ab5394ab23949 ,
                        0xaf3a8dd082fdfefdfffffffefefef6f6f6f9f9f9ffffffb3e0ac4bb63a4ab339 ,
                        0x49af3a4ab3394ab5394ab5394ab5394ab13a4ab23957ba476bc15d4ab2394ab4 ,
                        0x394ab5394ab5394ab53949b03a49b03a4ab5394ab5394ab5394ab43949af3a4a ,
                        0xb3394ab5394ab5394ab5394ab23949af3a8dd083fcfefcfffffffefefef6f6f6 ,
                        0xf9f9f9ffffffb2dfab4cb63b4ab33949af3a4ab3394ab5394ab5394ab5394ab3 ,
                        0x396dc15f87cd7c4ab4394ab03a4ab4394ab5394ab5394ab53949b03a49b03a4a ,
                        0xb5394ab5394ab5394ab43949af3a4ab3394ab5394ab5394ab5394ab23949af3a ,
                        0x8ed083fdfefdfffffffefefef6f6f6f9f9f9ffffffb3e0ac4bb53a4ab33949af ,
                        0x3a4ab3394ab5394ab5394ab53989ce7eb2dfab4ab5394ab33949af3a4ab4394a ,
                        0xb5394ab5394ab53949b03a49b03a4ab5394ab5394ab5394ab43949af3a4ab339 ,
                        0x4ab5394ab5394ab5394ab23949af3a8ed083fcfefcfffffffefefef6f6f6f9f9 ,
                        0xf9ffffffb2dfab4cb63b4ab33949af3a4ab3394ab5394ab539b6e1b0e1f2de4b ,
                        0xb53a4ab5394ab23949af3a4ab4394ab5394ab5394ab53949b03a49b03a4ab539 ,
                        0x4ab5394ab5394ab43949af3a4ab3394ab5394ab5394ab5394ab23949af3a8ed0 ,
                        0x83fdfefdfffffffefefef6f6f6f9f9f9fefffeb2e0ab4bb53a4ab33949af3a4a ,
                        0xb3394db63debf6e9fdfdfd6bc25d4ab5394ab5394ab23949af3a4ab4394ab539 ,
                        0x4ab5394ab53949b03a49b03a4ab5394ab5394ab5394ab43949af3a4ab3394ab5 ,
                        0x394ab5394ab5394ab23949af3a8ed083fcfefcfffffffefefef6f6f6f9f9f9ff ,
                        0xffff7fcb734ab5394ab33949af3a73c366fefffefdfdfdbbe2b54ab5394ab539 ,
                        0x4ab5394ab23949af3a4ab4394ab5394ab5394ab53949b03a49b03a4ab5394ab5 ,
                        0x394ab5394ab43949af3a4ab3394ab5394ab5394ab5394ab23949af3a8dd083fd ,
                        0xfefdfffffffefefef6f6f6e8f2e65dbd4e4ab5394ab5394ab339bcdfb6fefefe ,
                        0xfffffff4f7f462bc544ab5394ab5394ab5394ab23949af3a4ab4394ab5394ab5 ,
                        0x394ab53949b03a49b03a4ab5394ab5394ab5394ab43949af3a4ab3394ab5394a ,
                        0xb5394ab5394ab23949af3a8dd082fdfefcffffffecf7ea65bb5849b13a4ab539 ,
                        0x4ab53968c159fbfdfbfdfdfdfffffffefefec3e1bf49b1394ab5394ab5394ab5 ,
                        0x394ab23949af3a4ab4394ab5394ab5394ab53949b03a49b03a4ab5394ab5394a ,
                        0xb5394ab43949af3a4ab3394ab5394ab5394ab5394ab23949af3a88ce7ccdeac8 ,
                        0x64c0564ab53949af3a49b13a4bb53ad0ebcbfffffffefefefffffffffffffdfe ,
                        0xfd8ac9804ab2394ab5394ab5394ab5394ab23949af3a4ab4394ab5394ab5394a ,
                        0xb53949b03a49b03a4ab5394ab5394ab5394ab43949af3a4ab3394ab5394ab539 ,
                        0x4ab5394ab23949af3a4ab4394ab5394ab5394ab43949b03a8fcd85fefffeffff ,
                        0xfffffffffffffffffffffffffff5faf469bb5d4ab2394ab5394ab5394ab5394a ,
                        0xb23949af3a4ab4394ab5394ab5394ab53949b03a49b03a4ab5394ab5394ab539 ,
                        0x4ab43949af3a4ab3394ab5394ab5394ab5394ab23949af3a4ab4394ab5394ab5 ,
                        0x396ec361f0f3f0f8f8f8fffffffffffffefefeffffffffffffffffffeaf5e960 ,
                        0xb8534ab2394ab5394ab5394ab5394ab23949af3a4ab4394ab5394ab5394ab539 ,
                        0x49b03a49b03a4ab5394ab5394ab5394ab43949af3a4ab3394ab5394ab5394ab5 ,
                        0x394ab23949af3a4ab43965c057eef8edfefefef6f6f6fbfbfbfffffffdfdfdfb ,
                        0xfbfbfefefeffffffffffffeaf5e969bb5d4ab2394ab5394ab5394ab5394ab239 ,
                        0x49af3a4ab4394ab5394ab5394ab53949b03a49b03a4ab5394ab5394ab5394ab4 ,
                        0x3949af3a4ab3394ab5394ab5394ab5394ab2396dbd60eef7ecfffffffffffffe ,
                        0xfefef9f9f9fdfdfdfffffffbfbfbf7f7f7fefefefffffffffffff5faf48ac980 ,
                        0x49b1394ab5394ab5394ab5394ab23949af3a4ab4394ab5394ab5394ab53949b0 ,
                        0x3a49b03a4ab5394ab5394ab5394ab43949af3a4ab3394ab5394ab43990d186f1 ,
                        0xf6f0f4f4f4fefefefffffffffffffffffffdfdfdfffffffffffff8f8f8f7f7f7 ,
                        0xfefefefffffffffffffdfefdc3e1bf62bc544ab5394ab5394ab5394ab3394ab0 ,
                        0x3a4ab4394ab5394ab5394ab5394ab13a4ab1394ab5394ab5394ab5394ab4394a ,
                        0xb03a65be57cceac7fefefefffffff9f9f9f6f6f6fefefeffffffffffffffffff ,
                        0xfffffffffffffffffffbfbfbfbfbfbfffffffffffffffffffefefef4f7f4bbe2 ,
                        0xb56bc25d4bb53a4ab5394ab4394ab2394ab5394ab5394ab5394ab5394ab3394a ,
                        0xb3394ab5394bb53a6cc25fc0e4baf5f7f5fdfdfdfffffffffffffffffffbfbfb ,
                        0xfbfbfbfefefefffffffffffffffffffffffffffffffffffffdfdfdfefefeffff ,
                        0xfffffffffffffffffffffdfdfdfdfdfde1f2deb2dfab87cd7c6bc15d57b9474d ,
                        0xb63c4db53d56b9476cc25e89cd7eb4e0ade2f3e0fefffefffffffefefefdfdfd ,
                        0xfffffffffffffffffffffffffefefefefefeffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="Submit.bmp"

                    LayoutCachedLeft =10440
                    LayoutCachedTop =9590
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =9950
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =10620
                    Top =1160
                    Width =1436
                    Height =260
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txBaseID"
                    ControlSource ="BaseID"
                    Format ="General Number"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =10620
                    LayoutCachedTop =1160
                    LayoutCachedWidth =12056
                    LayoutCachedHeight =1420
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =10605
                            Top =900
                            Width =1455
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label24"
                            Caption ="BaseID"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =10605
                            LayoutCachedTop =900
                            LayoutCachedWidth =12060
                            LayoutCachedHeight =1155
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =2
                    OverlapFlags =255
                    Left =5175
                    Top =180
                    Width =7065
                    Height =10320
                    Name ="Box247"
                    GridlineColor =10921638
                    LayoutCachedLeft =5175
                    LayoutCachedTop =180
                    LayoutCachedWidth =12240
                    LayoutCachedHeight =10500
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =247
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =6900
                    Top =6460
                    Width =3660
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbFiller_2Pass"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateConstructionFiller"
                    ColumnWidths ="1440;1080;1080"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =6460
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =6775
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5400
                            Top =6460
                            Width =1260
                            Height =315
                            Name ="Label249"
                            Caption ="Filler"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =6460
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =6775
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    BorderWidth =3
                    Left =12780
                    Top =180
                    Width =9600
                    Height =10245
                    Name ="frmCreateConstruction_CableConstruction"
                    SourceObject ="Form.frmCreateConstruction_CableConstruction"

                    LayoutCachedLeft =12780
                    LayoutCachedTop =180
                    LayoutCachedWidth =22380
                    LayoutCachedHeight =10425
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =8700
                    Width =1500
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =24
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txFiberCount"
                    ControlSource ="FiberCount"
                    Format =">"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =8700
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =9015
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5400
                            Top =8700
                            Width =960
                            Height =315
                            Name ="Label256"
                            Caption ="Fiber Count"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =8700
                            LayoutCachedWidth =6360
                            LayoutCachedHeight =9015
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =5614
                    Width =1500
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txFibersPerBundle_1Pass"
                    ControlSource ="FibersPerBundle"
                    Format =">"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =5614
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =5929
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5400
                            Top =5610
                            Width =1320
                            Height =315
                            Name ="Label258"
                            Caption ="Fibers Per Bundle"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =5610
                            LayoutCachedWidth =6720
                            LayoutCachedHeight =5925
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =3840
                    Top =2640
                    Width =1230
                    Height =210
                    ForeColor =2366701
                    Name ="lbBaseExists_Err"
                    Caption ="* Base Exists"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =2640
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =2850
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =2520
                    Top =3000
                    Width =2580
                    Height =210
                    ForeColor =2366701
                    Name ="lbDesignCodeLen_Err"
                    Caption ="* Must have length of 7 characters"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =2520
                    LayoutCachedTop =3000
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =3210
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6900
                    Top =4760
                    Width =3660
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbCableType"
                    ControlSource ="CableType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT dbo_tblCableType.CableType FROM dbo_tblCableType ORDER BY dbo_tblCableTyp"
                        "e.CableType;"
                    ColumnWidths ="3600;1440;1440"
                    DefaultValue ="Null"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =4760
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =5075
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5400
                            Top =4755
                            Width =1020
                            Height =405
                            Name ="Label263"
                            Caption ="Cable Type"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =4755
                            LayoutCachedWidth =6420
                            LayoutCachedHeight =5160
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =10740
                    Top =5970
                    Width =426
                    Height =426
                    TabIndex =14
                    Name ="btnOpenSub_2pass"
                    Caption ="btnOpenSub"
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
                    ControlTipText ="Open Construction"
                    Picture ="OpenIcon.bmp"

                    LayoutCachedLeft =10740
                    LayoutCachedTop =5970
                    LayoutCachedWidth =11166
                    LayoutCachedHeight =6396
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =10740
                    Top =6450
                    Width =426
                    Height =426
                    TabIndex =17
                    Name ="btnOpenFiller_2Pass"
                    Caption ="btnOpenFiller"
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
                    ControlTipText ="Open Construction"
                    Picture ="OpenIcon.bmp"

                    LayoutCachedLeft =10740
                    LayoutCachedTop =6450
                    LayoutCachedWidth =11166
                    LayoutCachedHeight =6876
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =11220
                    Top =5970
                    Width =426
                    Height =426
                    TabIndex =15
                    Name ="btnClearSub_2Pass"
                    Caption ="btnOpenSub"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddddddd177ddddd77dd1ddd1177dddd17dd11dd7117ddd71ddd ,
                        0x111dd1177d117ddd1111d7117711dddd11111d11111ddddd1111dd71117ddddd ,
                        0x111d77111177dddd11d711dd71177ddd1dddddddd71177ddddddddddddd11ddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
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
                    ControlTipText ="Open Construction"

                    LayoutCachedLeft =11220
                    LayoutCachedTop =5970
                    LayoutCachedWidth =11646
                    LayoutCachedHeight =6396
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =11220
                    Top =6450
                    Width =426
                    Height =426
                    TabIndex =18
                    Name ="btnClearFiller_2Pass"
                    Caption ="btnOpenFiller"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddddddd177ddddd77dd1ddd1177dddd17dd11dd7117ddd71ddd ,
                        0x111dd1177d117ddd1111d7117711dddd11111d11111ddddd1111dd71117ddddd ,
                        0x111d77111177dddd11d711dd71177ddd1dddddddd71177ddddddddddddd11ddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
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
                    ControlTipText ="Open Construction"

                    LayoutCachedLeft =11220
                    LayoutCachedTop =6450
                    LayoutCachedWidth =11646
                    LayoutCachedHeight =6876
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    Left =5400
                    Top =2165
                    Width =1335
                    Height =315
                    Name ="Label275"
                    Caption ="Commoned Items"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =2165
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =2480
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6900
                    Top =2165
                    Width =1140
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =28
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="cbIsCommoned"
                    ControlSource ="IsCommoned"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;1440"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =2165
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =2480
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
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private blnGood, isConstructionLoaded As Boolean
Private CablePasses As Integer
Private CableCon As New clsCableConstruction
Public newCableCon As New clsCableConstruction
Private m_SubCon As New clsCableConstruction
Private m_SubFiberCount#, m_NumSubs#, m_FibersPerBundle#, m_NumFiller#, m_NumCuUnit#
Dim mlDesignCode As String

Private Property Get CableDesignCode() As String
mlDesignCode = Me.txDesignCode
CableDesignCode = mlDesignCode
End Property

Private Property Get SubDesignCodeOD() As String
SubDesignCodeOD = Right(SubCon.DesignCode, 2)
End Property

Private Property Get SubFiberCount() As Integer
m_SubFiberCount = Nz(Me.cbSubunit_2Pass.Column(1), 0)
SubFiberCount = m_SubFiberCount
End Property

Private Property Get NumSubCount() As Integer
m_NumSubs = Nz(Me.txNumSubs_2Pass, 0)
NumSubCount = m_NumSubs
End Property

Private Property Get SubCon() As clsCableConstruction
m_SubCon.Load (Me.cbSubunit_2Pass)
Set SubCon = m_SubCon
End Property

Private Property Get SubFibersPerBundle() As Integer
m_FibersPerBundle = SubCon.FibersPerBundle
SubFibersPerBundle = m_FibersPerBundle
End Property

Private Property Get NumFillerCount() As Integer
m_NumFiller = Nz(Me.txNumFillers_2Pass, 0)
NumFillerCount = m_NumFiller
End Property

Private Property Get NumCuUnitCount() As Integer
m_NumCuUnit = Nz(Me.txNumCu_2Pass, 0)
NumCuUnitCount = m_NumFiller
End Property

Private Sub TwoPassUpdate()
Me.txFibersPerBundle_1Pass = SubFibersPerBundle
Call Get2PassDesignCode
Call Get2PassDesignCodeExtension

Call updateCableIndicators

End Sub
Private Sub Get2PassDesignCode()
Me.txDesignCode = Left(Me.txDesignCode, 2) & Format(SubFiberCount * (NumSubCount - NumFillerCount - m_NumCuUnit), "000") & SubDesignCodeOD

End Sub

Private Sub Get2PassDesignCodeExtension()
Me.txDesignCodeExt = ":" & HexLetter(SubFiberCount) & HexLetter(NumSubCount) & HexLetter(SubFibersPerBundle)

End Sub


Private Sub btnAddNPR_Click()
DoCmd.OpenForm "frmCableFamily"
End Sub

Private Sub btnBurnListing_Click()

DoCmd.OpenForm "frmBurnListingAddNewBase"


Call Forms("frmBurnListingAddNewBase").btnAdd_Click

End Sub

Private Sub btnClearFiller_2Pass_Click()
Me.cbFiller_2Pass = Null
End Sub



Private Sub btnClearSub_2Pass_Click()
Me.cbSubunit_2Pass = Null
End Sub

Private Sub btnClose_Click()
If blnGood = False Then
    If MsgBox("Your work is not saved.  Are you sure you want to quit?", vbYesNo) = vbYes Then
        DoCmd.Close acForm, Me.name
    End If
Else
    DoCmd.Close acForm, Me.name
End If
End Sub
Private Sub ClearTables()

DoCmd.RunSQL "DELETE FROM tblCableConstructions_Local"
DoCmd.RunSQL "DELETE FROM tblCableConstructionReferences_Local"
End Sub

Private Sub btnCopy_Click()





On Error GoTo ErrorHandler:

If IsNull(Me.cbBase_Copy) Then
    Exit Sub
End If

Dim lResponse As String
isConstructionLoaded = False

Me.cbBase_Copy.SetFocus

CableCon.Load (Me.cbBase_Copy)



lResponse = MsgBox("Copy base and construction?", vbYesNo)


'Me.Requery
DoCmd.Echo False

Call ClearTables




DoCmd.OpenQuery "qryfrmCreateConstruction_Reference_LOAD"


'Check if users wants to load the construction
'User may only want to create a new base using an old construction
If lResponse = vbYes Then
    'Load construction information
    DoCmd.OpenQuery "qryfrmCreateConstruction_Construction_LOAD"
    isConstructionLoaded = True

End If


Me.Requery
Me.Refresh

CablePasses = Me.cbBase_Copy.Column(1)


Me.cbCablePasses = CablePasses

Me.cbSubunit_2Pass = CableCon.subBase
Me.cbFiller_2Pass = CableCon.fillerBase

mlDesignCode = Me.DesignCode

Call HideFields
Call updateCableIndicators
Call CheckBaseExists
Call CheckStdOpExists
DoCmd.Echo True

ErrorExit:
Exit Sub

ErrorHandler:
If Err.Number = 3167 Then
    Resume
Else
MsgBox (Err.Number & " - " & Err.Description), vbCritical
End If
Resume ErrorExit




End Sub

Private Sub btnOpenBase_Click()
    If IsNull(Me.cbBase_Copy) Then
        Exit Sub
    End If
    
    Call OpenAClient(3, Me.cbBase_Copy, Nz(Me.cbBase_Copy.Column(1), 2))
End Sub

Private Sub btnOpenFiller_2Pass_Click()
    If IsNull(Me.cbFiller_2Pass) Then
        Exit Sub
    End If

    Call OpenAClient(3, Me.cbFiller_2Pass, 1)
End Sub

Private Sub btnOpenSub_2pass_Click()

    If IsNull(Me.cbSubunit_2Pass) Then
        Exit Sub
    End If
Call OpenAClient(3, Me.cbSubunit_2Pass, 1)
End Sub


Private Sub btnSave_Click()
Dim StdOp As String
Dim Base As String
Dim CableCheck As New clsCableConstruction
Dim ctl As Variant
Dim errMsg As String

On Error GoTo ErrorHandler:
DoCmd.SetWarnings False

    blnGood = True
    Base = Me.txBase
    
    errMsg = "Red issues must be resolved before continuing"
    For Each ctl In Me.Controls
        If ctl.name Like "*_Err" And ctl.Visible Then
            MsgBox errMsg, vbCritical
            Exit Sub
        End If
    Next ctl
    
    If isConstructionLoaded Then
        For Each ctl In Me.frmCreateConstruction_CableConstruction.Form.Controls
        If ctl.name Like "*_Err" And ctl.Visible Then
            MsgBox errMsg, vbCritical
            Exit Sub
        End If
        Next ctl
    End If
    
    Me.Requery
    Me.Refresh
    Me.btnCopy.SetFocus

    
    'Check to see if construction is loaded
    If isConstructionLoaded Then
        StdOp = Me.frmCreateConstruction_CableConstruction.Form.txSetup
        'Load cable construction
        DoCmd.OpenQuery "qryNewCableConstructionLOAD"
        'Check to see if the cable loaded correctly
        If CableCheck.SetupExist(StdOp) = False Then
            MsgBox "Cable load unsuccessful.  New construction did not save", vbCritical
            Exit Sub
        End If
        'Get BaseID to load into the construction References
        Me.txBaseID = CableCheck.BaseID
            Me.Requery
    End If
    'Load Cable Construction Reference data into live table
    
        DoCmd.OpenQuery "qryNewCableConstructionReferenceLOAD"
    
    
    'Check to see if construction references loaded correctly
    If CableCheck.BaseExist(Base) = False Then
        MsgBox "Cable load unsuccessful.", vbCritical
        Exit Sub
    End If
    
    
    If Me.cbCablePasses = 2 Then
    'Load sub
        If IsNull(Me.cbSubunit_2Pass) = False Then
            DoCmd.RunSQL ("INSERT INTO [Product Desrcriptions]( [Level 1 Product], ProductID) VALUES ( '" & Me.cbSubunit_2Pass & "'," & CableCheck.SetupID & ")")
        End If
    'Load Filler
        If IsNull(Me.cbFiller_2Pass) = False Then
            DoCmd.RunSQL ("INSERT INTO [Product Desrcriptions]([Level 1 Product], ProductID, SubFiller) VALUES ( '" & Me.cbFiller_2Pass & "'," & CableCheck.SetupID & ", 1)")
        End If
    End If
    
    Beep
    MsgBox "Construction has been loaded"
    
    If MsgBox("Would you like to create another construction?", vbYesNo) = vbYes Then
        Call ClearTables
        Me.Requery
        Me.Refresh
    Else
    
        DoCmd.Close acForm, Me.name
    End If
    
ErrorExit:
Exit Sub

ErrorHandler:
MsgBox (Err.Number & " - " & Err.Description), vbCritical
Resume ErrorExit


End Sub

Private Sub cbBase_Copy_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = vbKeyReturn Then
        Me.btnCopy.SetFocus
        Call btnCopy_Click
    End If

End Sub

Private Sub cbDesignGroup_AfterUpdate()
Me.cbDesignType = Null
Me.cbDesignType.Requery
Call cbDesignType_AfterUpdate

End Sub

Private Sub cbCablePasses_AfterUpdate()
Call HideFields
End Sub

Private Sub HideFields()

Dim bVisible As Boolean
Dim ctl As control

If Me.cbCablePasses = 1 Then
    bVisible = False
Else
    bVisible = True
End If
    

For Each ctl In Me.Controls
    If ctl.name Like "*2pass" Then
        ctl.Visible = bVisible
    'ElseIf ctl.Name Like "*1pass" Then
        'ctl.Visible = Not bVisible
    End If
Next ctl

If isConstructionLoaded = True Then
    Me.frmCreateConstruction_CableConstruction.Form.txCablePasses = Me.cbCablePasses
End If

End Sub

Private Sub cbDesignType_AfterUpdate()
Me.cbFamilyID = Null
Me.cbFamilyID.Requery

End Sub


Private Sub cbFiller_2Pass_AfterUpdate()
Call TwoPassUpdate
End Sub

Private Sub cbSubunit_2Pass_AfterUpdate()
Call TwoPassUpdate
End Sub

Private Sub Form_Delete(Cancel As Integer)
Me.Requery
Me.Refresh
End Sub

Private Sub Form_Load()
    Call ClearTables
    Me.Requery
    Me.Refresh
End Sub

Private Sub txBase_AfterUpdate()
Dim Exist As Boolean

Exist = CheckBaseExists()
If Exist = False Then
    newCableCon.Base = Me.txBase
End If

End Sub

Private Sub txDesignCode_AfterUpdate()

Call updateCableIndicators

End Sub

Private Sub txDesignCode_Change()
mlDesignCode = Me.txDesignCode.Text
Call updateCableIndicators

End Sub

Private Sub txDesignCodeExt_AfterUpdate()
Call updateCableIndicators
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Const conDuplicateKey = 3146
    Const conKey = 3621
    Dim strErrMsg As String

    Select Case DataErr
        Case conDuplicateKey:
            strErrMsg = "Record already exists in the database!"
        Case conKey:
            strErrMsg = "Key field must be completed to save"
        Case Else
            Response = acDataErrDisplay
    End Select

    If strErrMsg <> vbNullString Then
        MsgBox strErrMsg, vbCritical, Me.Caption
        ActiveControl.Undo
        Response = acDataErrContinue
    End If

End Sub


Private Sub updateCableIndicators()
Dim isLen As Boolean

Dim FiberBundle, iFiberCount As Integer
Dim FiberCount As String


isLen = Len(CableDesignCode) <> 7

Call PaintExistingRed(Me.txDesignCode, Me.lbDesignCodeLen_Err, isLen)

Me.txCableFamily = Left(CableDesignCode, 1)
Me.txCableLevel1 = Mid(CableDesignCode, 2, 1)
FiberCount = Mid(CableDesignCode, 3, 3)

If IsNumeric(FiberCount) Then
    FiberCount = FiberCount
Else
    FiberCount = 0
End If

iFiberCount = FiberCount

'Commented out logic
'If a single pass cable then hide the fibers per bundle field
'If Me.cbCablePasses = 1 Then
    If Me.txDesignCodeExt Like "*:*" Then
    'Logic for determining fiber bundle from design code extension
    ElseIf Me.cbTBType.Column(1) = 3 Or Me.cbTBType = 31 Or Me.cbTBType = 9 Then
        If iFiberCount = 0 Then
            FiberBundle = 0
        ElseIf iFiberCount Mod 12 = 0 Then
            FiberBundle = 12
        ElseIf iFiberCount Mod 8 = 0 Then
            FiberBundle = 8
        ElseIf iFiberCount Mod 9 = 0 Then
            FiberBundle = 9
        Else
            FiberBundle = FiberCount
        End If
    ElseIf Me.cbTBType = 20 Then
        FiberBundle = 4
            Me.txFibersPerBundle_1Pass = FiberBundle
    ElseIf iFiberCount > 0 Then
        Me.txFibersPerBundle_1Pass = 1
    Else
        Me.txFibersPerBundle_1Pass = Null
    End If
'Else
'    Me.txFibersPerBundle_1Pass = Null
'End If

Me.Base = Left(Me.txDesignCode, 5) & "-" & Mid(Me.txDesignCode, 6, 1) & "." & Mid(Me.txDesignCode, 7, 1) & Nz(Me.DesignExtension, "")

Call txBase_AfterUpdate

Me.txFiberCount = FiberCount

End Sub

Private Function CheckBaseExists() As Boolean
Dim Cable As New clsCableConstruction
Dim Exists As Boolean

Exists = Cable.BaseExist(Base)

Call PaintExistingRed(Me.txBase, Me.lbBaseExists_Err, Exists)


End Function

Private Sub CheckStdOpExists()

'Paint the background of control and show label
If isConstructionLoaded Then
    Call Me.frmCreateConstruction_CableConstruction.Form.txSetup_AfterUpdate
End If

End Sub

Public Sub PaintExistingRed(TextBox As control, Warning As control, Exists As Boolean)
'Paint the background of control and show label
If Exists = True Then
    TextBox.BackColor = RGB(255, 0, 0)
    Warning.Visible = True
Else
    TextBox.BackColor = RGB(255, 255, 255)
    Warning.Visible = False
End If


End Sub

Public Sub SetIsConstructionLoaded(IsLoaded As Boolean)

isConstructionLoaded = IsLoaded


End Sub

Private Sub txNumFillers_2Pass_AfterUpdate()
Call TwoPassUpdate
End Sub

Private Sub txNumSubs_2Pass_AfterUpdate()
Call TwoPassUpdate
End Sub

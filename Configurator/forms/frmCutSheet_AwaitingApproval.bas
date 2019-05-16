Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10695
    DatasheetFontHeight =10
    ItemSuffix =77
    Left =10830
    Top =4425
    Right =21525
    Bottom =14310
    DatasheetGridlinesColor =12632256
    Filter ="[Base] like \"*\""
    OrderBy ="[tblCutSheetApproval_Local].[RequestedDate] DESC, [tblCutSheetApproval_Local].[B"
        "ase]"
    RecSrcDt = Begin
        0x391125016ff9e440
    End
    RecordSource ="tblCutSheetApproval_Local"
    Caption ="Item Search"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =3
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =90.0
            ShowPageHeaderAndPageFooter =1
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
                    Left =3660
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Base"
                    LayoutCachedLeft =3660
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =1680
                    Top =1680
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Oracle Part"
                    LayoutCachedLeft =1680
                    LayoutCachedTop =1680
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =2520
                    Top =60
                    Width =5640
                    Name ="Box35"
                    LayoutCachedLeft =2520
                    LayoutCachedTop =60
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =780
                End
                Begin Label
                    OverlapFlags =85
                    Left =8340
                    Top =1680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Approve"
                    LayoutCachedLeft =8340
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =120
                    Top =1440
                    Width =10500
                    Name ="Line58"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =1440
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7320
                    Top =900
                    Width =479
                    TabIndex =2
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddd99dddddd9ddddddd99ddddd9ddddddddd99ddd99d ,
                        0xddddddddd99999ddddddd777dd999dddddddd78dd99999ddddddd7d9999dd999 ,
                        0xddddd7d999ddddddddddd78ddddddddddddd87778dddddddddd8777778dddddd ,
                        0xddd77777778ddddddd777777777dddddd77777777777ddddd77777777777dddd ,
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

                    LayoutCachedLeft =7320
                    LayoutCachedTop =900
                    LayoutCachedWidth =7799
                    LayoutCachedHeight =1260
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =2640
                    Top =180
                    Width =5400
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblApproval"
                    Caption ="Cut Sheet Technical Approval"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =180
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =675
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =900
                    Width =479
                    TabIndex =1
                    Name ="btnFilter"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddd0000dddddddddddd0660dddddddddddd0660dddddddddddd0f80dddddd ,
                        0xdddddd0f80ddddddddddd088860ddddddddd06888660ddddddd068f888660ddd ,
                        0xdd068f88888660ddd068fff88886660dd00000000000000ddddddddddddddddd ,
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

                    LayoutCachedLeft =6660
                    LayoutCachedTop =900
                    LayoutCachedWidth =7139
                    LayoutCachedHeight =1260
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
                    Left =2820
                    Top =900
                    Width =3600
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =900
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =85
                    Left =5580
                    Top =1680
                    Width =2700
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label73"
                    Caption ="Item No"
                    LayoutCachedLeft =5580
                    LayoutCachedTop =1680
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =1680
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label75"
                    Caption ="Request Date"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =9360
                    Top =180
                    Width =720
                    Height =660
                    TabIndex =3
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
                    ControlTipText ="Close Form"
                    Picture ="CloseIcon.bmp"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =180
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3660
                    Top =60
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txBase"
                    ControlSource ="Base"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =60
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txOracle"
                    ControlSource ="Oracle"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8640
                    Top =120
                    Width =240
                    TabIndex =5
                    BorderColor =15916467
                    Name ="chApproved"
                    ControlSource ="Technical_Approval"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =8640
                    LayoutCachedTop =120
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =60
                    Width =2700
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="Text74"
                    ControlSource ="ItemNo"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =60
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9180
                    Width =540
                    Name ="btnCutSheet"
                    Caption ="Cut Sheet"
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
                        0xfffffffffffffffffffffffffefefefcfcfcfcfcfcfcfcfcfcfcfcfcfcfcfdfd ,
                        0xfdfdfdfdfefefefefefeffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffdfdfdeeeeeee6e6 ,
                        0xe6e6e6e6e6e6e6e7e7e7e7e7e7e9e9e9ebebebecececeeeeeeefefefefefeff1 ,
                        0xf1f1f2f2f2f3f3f3f5f5f5f8f8f8fbfbfbfdfdfdffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffff6f6f6d9d9d9d6d6d6d6d6d6d5d5d5d5d5d5d4d4d4d4d4d4d4 ,
                        0xd4d4d3d3d3d4d4d4d5d5d5d5d5d5d6d6d6d7d7d7dadadadcdcdcdfdfdfe1e1e1 ,
                        0xe4e4e4e9e9e9efefeff4f4f4fbfbfbfefefeffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffe6e6e6d8d8d8d8d8d8d7 ,
                        0xd7d7d7d7d7d6d6d6d6d6d6d5d5d5d5d5d5d5d5d5d4d4d4d4d4d4d4d4d4d3d3d3 ,
                        0xd2d2d2d2d2d2d1d1d1d1d1d1d0d0d0d1d1d1d3d3d3d4d4d4d4d4d4dcdcdce7e7 ,
                        0xe7f4f4f4fdfdfdffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffff0f0f0dadadad9d9d9d9d9d9d9d9d9d8d8d8d8d8d8d8d8d8d7d7d7d7d7d7 ,
                        0xd6d6d6d5d5d5d5d5d5d4d4d4d4d4d4d3d3d3d3d3d3d3d3d3d2d2d2d2d2d2d1d1 ,
                        0xd1d1d1d1d0d0d0d0d0d0d0d0d0d1d1d1eaeaeafefefeffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffff4f4f4dcdcdcdbdbdbdadadadadadadadada ,
                        0xd9d9d9d9d9d9d8d8d8d8d8d8d7d7d7d7d7d7d7d7d7d6d6d6d6d6d6d5d5d5d5d5 ,
                        0xd5d5d5d5d4d4d4d3d3d3d3d3d3d2d2d2d2d2d2d1d1d1d1d1d1d0d0d0dadadaf7 ,
                        0xf7f7fffffffffffffffffffffffffffffffffffffffffffffffff6f6f6dddddd ,
                        0xdddddddcdcdcdcdcdcdbdbdbdadadadadadadadadad9d9d9d9d9d9d8d8d8d8d8 ,
                        0xd8d8d8d8d7d7d7d7d7d7d6d6d6d6d6d6d5d5d5d5d5d5d5d5d5d4d4d4d4d4d4d3 ,
                        0xd3d3d3d3d3d2d2d2d3d3d3eeeeeefefefeffffffffffffffffffffffffffffff ,
                        0xfffffffffffff7f7f7dfdfdfdfdfdfdededededededddddddddddddcdcdcdcdc ,
                        0xdcdbdbdbdbdbdbdadadadadadadadadad9d9d9d9d9d9d8d8d8d8d8d8d7d7d7d7 ,
                        0xd7d7d7d7d7d6d6d6d6d6d6d5d5d5d5d5d5d4d4d4d4d4d4e1e1e1fcfcfcffffff ,
                        0xfffffffffffffffffffffffffffffffffffff7f7f7e1e1e1e0e0e0e0e0e0e0e0 ,
                        0xe0dfdfdfdfdfdfdedededddddddddddddddddddcdcdcdcdcdcdbdbdbdbdbdbda ,
                        0xdadadadadadadadad9d9d9d9d9d9d8d8d8d8d8d8d8d8d8d7d7d7d7d7d7d6d6d6 ,
                        0xd6d6d6d8d8d8fbfbfbfffffffdfdfedcdde9bebed6bebed7bebed7bfbfd7b9b9 ,
                        0xd1adaec6aeaec6aeafc6adaec5aeafc6aeb0c6afb0c6b0b1c7b2b1c8b2b3c8b3 ,
                        0xb4c7b5b4c7b5b5c7b6b6c7b6b6c7b6b6c7b7b6c8b7b7c8b8b8c7b8b8c7c2c2cc ,
                        0xd8d8d9d9d9d9d9d9d9d8d8d8d7d7d7d7d7d7f6f6f6fffffff8f8fc7e7dc70301 ,
                        0x9503019503019504029503029603019702019702019702019802019802019704 ,
                        0x0196030196030396030396030297030196030195030295030195030196030296 ,
                        0x03019503019503038d3e3f92ceced4dadadadadadadadadad9d9d9d9d9d9f1f1 ,
                        0xf1fffffff9f9fe7f7fe90101d60101d60a0ad91818c51212c50201d90100dc00 ,
                        0x00dc0000de0304df1919cf1918c91918ce1616d70e0fde0303e30001e7090ae3 ,
                        0x1716d30d0dd10102e90000ea0000ed0000ed0201d93d3ea0cfcfd6dcdcdcdcdc ,
                        0xdcdcdcdcdbdbdbd9d9d9efefeffffffff9f9fe7f7fe60101cf0101cf4546dbe1 ,
                        0xe2f39d9ed70201cd0100d40100d50000d71b1adbe1e1efdcdbecd0d0e8cccbe3 ,
                        0x9797ce3332c30403da4949e5d1d2f17373cf0a09e20100e40000e60001e70201 ,
                        0xd43e3ea0d1d1d7dddddddddddddddddddcdcdcdbdbdbeeeeeefffffff9f9fe7f ,
                        0x7fe20101c80100c94e4ed8fbfbfeb6b7e1191ac40f10c40303c80000d01d1cd6 ,
                        0xf9f9fbbdbce86b6bdc9595e9ededf8d6d6e72525bf5151e4e9e9fb7f7fd30a09 ,
                        0xdb0000dd0101df0000e10101cf3f3ea0d3d3d9dfdfdfdfdfdfdfdfdfdedededd ,
                        0xddddf0f0f0fffffff9f9fd7e7fde0100c10100c24d4ed3fbfbfee6e5f2b0b0e1 ,
                        0xa3a4dc6e6fcb0d0dbd1d1ccff9f9fa9797db0707c50c0ccfabacedfbfbfd7070 ,
                        0xc75151dfe9e9fab3b3e16d6ed86768d94142ce0101d80101c93e3f9fd4d4dbe1 ,
                        0xe1e1e1e1e1e1e1e1e0e0e0dfdfdff2f2f2fffffff9f9fd7e7fdb0101b90100ba ,
                        0x4c4ecefbfbfec4c4e45555cdbcbbebf1f1f75354ba1f1ec8f8f9fa9797d90607 ,
                        0xbf0202c79090e4fcfbfc8182d05051dae9e9faebecf5d8daf2d6d9f38688da02 ,
                        0x01d20101c33e3f9ed6d7dde3e3e3e3e3e3e3e3e3e2e2e2e1e1e1f5f5f5ffffff ,
                        0xf9f9fd7f7fd80001b20100b34d4ec9fbfbfeb0b0dd0909b4898adbf6f6fc7576 ,
                        0xc82221c2f9f9fa9898d70d0db92d2dc2dddef2f9f9fa4446ca5250d7e9e9f983 ,
                        0x82cf1110c60908c70908c80201cc0101bd3f3f9ed8d8dfe5e5e5e5e5e5e5e5e5 ,
                        0xe3e3e3e4e4e4fafafafffffff9f9fd7e7fd50001ac0101ad4d4dc3f7f6fbe4e4 ,
                        0xf0babadbe1e1eee6e7f63737c0201fbdf6f6f9e0e0efc5c5e1e3e5f0eceaf596 ,
                        0x99dd0605bd4f4fd0e4e4f6dbdcecc0c2e0bebfde9295ce0303bb0201b740409d ,
                        0xdadae0e6e6e6e6e6e6e6e6e6e5e5e5e6e6e6fbfbfbfffffff9f9fd7e7fd10102 ,
                        0xa60102a73334b4a3a4dda6a6dea4a4dd8f90d63e40bf0607b11415b3a3a4dda5 ,
                        0xa6df9e9edf8788d74848c60708b50000b73635c49898dea6a6e3a6a6e4a5a6e4 ,
                        0x8181d60303bb0201b340409cdbdbe1e8e8e8e8e8e8e8e8e8e7e7e7e8e8e8fcfc ,
                        0xfcfffffff9f9fd8788d01313a51312a61311a61210a71211a81210a9110faa11 ,
                        0x10ac1010ac1010ad1211ae1211af1210b01010b01010b21010b3100fb31110b5 ,
                        0x1111b51111b61110b81111ba1111ba1010bb100fb14f4fa2dedee4eaeaeaeaea ,
                        0xeae9e9e9e8e8e8ecececfdfdfdfffffffdfdfed2d3eaa9a9dbaaaadca7a6d9a2 ,
                        0xa0d5a1a0d6a1a0d5a1a1d6a2a1d7a1a2d7a1a2d7a1a2d8a1a2d8a1a2d8a0a1d8 ,
                        0xa0a1d9a0a0da9fa0d99fa0da9fa0da9f9fd99f9fda9e9fdb9f9fda9e9fda9e9e ,
                        0xd8b7b7d4e8e8eaececececececebebebebebebefefeffefefefffffffffffffe ,
                        0xfffefefffefffffffaf9f9f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2 ,
                        0xf2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2 ,
                        0xf2f1f1f1f0f0f0efefefefefefefefeeefefefeeeeeeededededededecececf2 ,
                        0xf2f2fefefefffffffffffffffffffffffffffffffbfbfbf2f2f2f2f2f2f2f2f2 ,
                        0xf2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2 ,
                        0xf2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f1f1f1f1f1f1f0f0f0f0 ,
                        0xf0f0efefefefefefeeeeeef5f5f5fefefeffffffffffffffffffffffffffffff ,
                        0xfdfdfdf4f4f4f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f0f0f0eded ,
                        0xedf0f0f0f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2 ,
                        0xf2f2f2f2f2f2f2f2f2f2f2f1f1f1f1f1f1f1f1f1f0f0f0f7f7f7fefefeffffff ,
                        0xfffffffffffffffffffffffffffffff9f9f9f3f3f3f2f2f2f2f2f2f2f2f2f2f2 ,
                        0xf2f2f2f2efefefd6d6d6bebebed1d1d1e2e2e2e9e9e9f0f0f0f2f2f2f2f2f2f2 ,
                        0xf2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2 ,
                        0xf2f2f2f9f9f9fffffffffffffffffffffffffffffffffffffffffffefefef8f8 ,
                        0xf8f3f3f3f2f2f2f1f1f1f0f0f0ebebebdfdfdfd0d0d0ddddddcfcfcfc5c5c5ca ,
                        0xcacad8d8d8e1e1e1ebebebf1f1f1f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2 ,
                        0xf2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f9f9f9ffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffefefefbfbfbf7f7f7f6f6f6f4f4f4f0f0f0efefeff9 ,
                        0xf9f9fefefefefefef8f8f8ebebebe1e1e1dbdbdbdadadadbdbdbe2e2e2eaeaea ,
                        0xf0f0f0f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f2f9f9f9ffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffefefefd ,
                        0xfdfdfefefefffffffffffffffffffffffffffffffffffffffffffefefefcfcfc ,
                        0xf9f9f9f5f5f5f0f0f0e9e9e9e3e3e3e5e5e5ebebebf0f0f0f2f2f2f2f2f2f2f2 ,
                        0xf2f2f2f2f2f2f2f8f8f8ffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffefefefefefefcfcfcf7f7 ,
                        0xf7f1f1f1efefeff0f0f0f0f0f0f2f2f2f2f2f2f8f8f8fefefeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffefefefbfbfbf9f9f9f7f7f7f9 ,
                        0xf9f9fefefeffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffe ,
                        0xfefeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
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
                    Picture ="download.bmp"

                    LayoutCachedLeft =9180
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9960
                    Width =420
                    Height =420
                    TabIndex =1
                    Name ="btnFind"
                    Caption ="Construction"
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

                    LayoutCachedLeft =9960
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1380
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text76"
                    ControlSource ="RequestedDate"
                    Format ="Short Date"
                    ShowDatePicker =0

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =780
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =8040
                    Top =240
                    Width =1320
                    Name ="btnApprove"
                    Caption ="Approve"
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
                    ControlTipText ="Approve Cut Sheets"
                    Picture ="Submit.bmp"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =240
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =600
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Top =480
                    Width =4380
                    Height =240
                    BorderColor =-2147483633
                    ForeColor =4144959
                    Name ="lbMessage"
                    Caption ="| Type to search  | Enter to select  | Esc to cancel"
                    LayoutCachedTop =480
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =720
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
Dim iApprovedClick As Integer
Dim currentUser As New clsUser


Private Sub btnApprove_Click()
Dim sEmail As String
Dim rst As New ADODB.Recordset

On Error GoTo ErrorHandler


If iApprovedClick <= 0 Then
MsgBox "No designs have been selected", vbCritical
    Exit Sub
End If


If MsgBox("Are you ready to approve these designs?", vbYesNo) = vbNo Then
   Exit Sub
End If

Me.txOracle.SetFocus
Me.Requery

DoCmd.SetWarnings False


DoCmd.RunSQL "DELETE FROM tblCutSheet_QueuedForApproval"

DoCmd.OpenQuery "qryCutSheet_QueuedForApproval_LOAD"


If iCutSheetApprover = 1 Then

    Call CutSheetApproval("", 4)

ElseIf iCutSheetApprover = 2 Then

    Call CutSheetApproval("", 5)

End If

'sBases = Split(BaseQueuedForApproval, ";") 'Fetch all approved bases and input into a semicolon delimited array
'For i = 0 To UBound(sBases)
'    If CutSheet_CheckApproval(sBases(i), sEmail) = True Then 'Fetch Email of requestor and determine if the Technical and Commerical review have been completed
'        Call CutSheetEmailApproved(sBases(i), sEmail) ' If both reviews have been completed then email the requestor
'    End If
'Next i

Set rst = BaseQueuedForApproval


Do Until rst.EOF
    If CutSheet_CheckApproval(rst!Base, sEmail) = True Then 'Fetch Email of requestor and determine if the Technical and Commerical review have been completed
        Call CutSheetEmailApproved(rst!Base, sEmail, rst!Oracle) ' If both reviews have been completed then email the requestor
    End If
    rst.MoveNext

Loop



BaseQueuedForApproval.Close

Beep
MsgBox "Approval Complete"

Form_frmFindCutSheet.Requery

Call Form_Open(0)

ErrorExit:
Exit Sub

ErrorHandler:
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit

End Sub

Private Sub btnClose_Click()
DoCmd.Close acForm, Me.name
End Sub



Private Sub btnCutSheet_Click()
If IsNull(Me.Oracle) Then
    Exit Sub
End If

OpenCutSheet (Me.Oracle)
End Sub

Private Sub btnFilter_Click()

Dim strFilter As String
strFilter = "[Oracle] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Base] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [ItemNo] like """ & Me.ItemNoSearch & "*"""

Me.Filter = strFilter
Me.FilterOn = True
Call Refresh_Click
End Sub



Private Sub btnFind_Click()

Call OpenAClient(InstanceFormType.ItemOracle, Me.Oracle, 2)
End Sub

Private Sub btnRemoveFilter_Click()
Me.ItemNoSearch.value = Null
Me.ItemNoSearch.Requery
Me.Filter = ""
Me.FilterOn = False
Call Refresh_Click
End Sub

Private Sub chApproved_AfterUpdate()

If Me.chApproved = 0 Then
    iApprovedClick = iApprovedClick - 1
Else
    iApprovedClick = iApprovedClick + 1
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
Dim strFilter As String
strFilter = "[Base] like """ & Me.ItemNoSearch & "*"""

Me.Filter = strFilter
Me.FilterOn = True

DoCmd.SetWarnings False

DoCmd.RunSQL "DELETE FROM tblCutSheetApproval_Local"

If iCutSheetApprover = 1 Then
    Me.lblApproval.Caption = "Cut Sheet Technical Approval"
    DoCmd.OpenQuery "qryCutSheetRequested_Tech_LOAD"

ElseIf iCutSheetApprover = 2 Then
    Me.lblApproval.Caption = "Cut Sheet Commercial Approval"
    DoCmd.OpenQuery "qryCutSheetRequested_Commerical_LOAD"

End If

Me.Requery

currentUser.Load

Me.ItemNoSearch.SetFocus


End Sub

Sub Refresh_Click()

Me.Refresh
Me.Requery
End Sub

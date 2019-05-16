Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =27960
    DatasheetFontHeight =10
    ItemSuffix =90
    Left =8175
    Top =6105
    Right =-29386
    Bottom =20205
    DatasheetGridlinesColor =12632256
    Filter ="([ItemNo] like \"PR00224*\" OR [Oracle] like \"PR00224*\" OR [Customer] like \"P"
        "R00224*\" OR [Fiber] like \"PR00224*\" OR [Fiber2] like \"PR00224*\" OR [JacketM"
        "aterial] like \"PR00224*\" OR [StandardOperation] like \"PR00224*\" OR [base] li"
        "ke \"PR00224*\" OR [Color] like \"PR00224*\" OR [DesignCode] like \"PR00224*\" O"
        "R [PID] like \"PR00224*\")"
    OrderBy ="[CreationDate] DESC"
    RecSrcDt = Begin
        0x343f87489410e540
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
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CustomControl
            OldBorderStyle =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            ForeThemeColorIndex =2
            ForeShade =50.0
        End
        Begin Page
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
                    Left =24360
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text8"
                    Caption ="Color"
                    LayoutCachedLeft =24360
                    LayoutCachedTop =1680
                    LayoutCachedWidth =25620
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =22440
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text10"
                    Caption ="Base"
                    LayoutCachedLeft =22440
                    LayoutCachedTop =1680
                    LayoutCachedWidth =24300
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =16500
                    Top =1680
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label15"
                    Caption ="Fiber 1"
                    LayoutCachedLeft =16500
                    LayoutCachedTop =1680
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =10260
                    Top =1680
                    Width =2940
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Customer"
                    LayoutCachedLeft =10260
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =25680
                    Top =1680
                    Width =600
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label19"
                    Caption ="UOM"
                    LayoutCachedLeft =25680
                    LayoutCachedTop =1680
                    LayoutCachedWidth =26280
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
                    Caption ="Active"
                    LayoutCachedLeft =7620
                    LayoutCachedTop =1680
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =17940
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label33"
                    Caption ="Fiber 2"
                    LayoutCachedLeft =17940
                    LayoutCachedTop =1680
                    LayoutCachedWidth =19200
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
                    Left =13260
                    Top =1680
                    Width =2220
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label39"
                    Caption ="PID"
                    LayoutCachedLeft =13260
                    LayoutCachedTop =1680
                    LayoutCachedWidth =15480
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =15540
                    Top =1680
                    Width =900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="Spec Rev"
                    LayoutCachedLeft =15540
                    LayoutCachedTop =1680
                    LayoutCachedWidth =16440
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =19260
                    Top =1680
                    Width =1200
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label43"
                    Caption ="Jacket"
                    LayoutCachedLeft =19260
                    LayoutCachedTop =1680
                    LayoutCachedWidth =20460
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =26340
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label45"
                    Caption ="Passes"
                    LayoutCachedLeft =26340
                    LayoutCachedTop =1680
                    LayoutCachedWidth =27000
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
                    Caption ="Reason"
                    LayoutCachedLeft =8340
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10200
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
                    Left =7680
                    Top =780
                    Width =480
                    TabIndex =9
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
                    ControlTipText ="Remove Filter"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =780
                    LayoutCachedWidth =8160
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
                    Left =21300
                    Top =1680
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label61"
                    Caption ="Std Op"
                    LayoutCachedLeft =21300
                    LayoutCachedTop =1680
                    LayoutCachedWidth =22380
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =7020
                    Top =780
                    Width =480
                    TabIndex =8
                    Name ="btnSearch"
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
                    ControlTipText ="Apply Filter"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =780
                    LayoutCachedWidth =7500
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
                    TabIndex =7
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3180
                    LayoutCachedTop =780
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =6240
                    Top =1320
                    Width =480
                    TabIndex =15
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

                    LayoutCachedLeft =6240
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6720
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
                    Left =8340
                    Top =780
                    Width =480
                    TabIndex =10
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

                    LayoutCachedLeft =8340
                    LayoutCachedTop =780
                    LayoutCachedWidth =8820
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
                    TabIndex =12
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
                    Width =780
                    Height =660
                    Name ="Weight"
                    Caption ="Item Weight"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000002400000024000000010008000000000010050000c40e0000c40e0000 ,
                        0x000100000001000012151aff13161bff13161cff14171cff15181dff16191eff ,
                        0x171a1fff181b20ff191c21ff1a1d22ff1b1e22ff1c1f24ff1d1f24ff1d2025ff ,
                        0x1e2025ff1f2227ff202227ff212328ff212429ff22252aff23262bff24262bff ,
                        0x24272cff25282dff292b30ff292c31ff2b2d32ff2b2e33ff2c2f33ff2c2f34ff ,
                        0x2d3034ff2f3237ff303237ff313338ff323439ff33353aff34363bff35373bff ,
                        0x35373cff35383cff36393dff37393eff373a3eff383b3fff393c40ff3a3d41ff ,
                        0x3c3e43ff3e4044ff3f4145ff404347ff414448ff424448ff44464aff45474bff ,
                        0x46484cff494b4fff4a4c50ff4b4d51ff4c4e52ff4d4f53ff515357ff525458ff ,
                        0x55575bff585a5eff595b5fff5a5c60ff5b5d61ff5c5e61ff5d5f62ff5d6063ff ,
                        0x5e6064ff606265ff616366ff626467ff646669ff66686bff686a6dff696b6eff ,
                        0x6e7073ff707175ff717376ff747679ff75777aff77797cff797a7dff7c7e80ff ,
                        0x7d7f82ff7e8083ff828386ff828487ff838487ff848588ff848688ff858689ff ,
                        0x86888bff888a8cff898b8dff8a8c8eff8b8c8fff8d8e91ff8f9092ff909193ff ,
                        0x909294ff919395ff939497ff949598ff949698ff959699ff98999bff999a9dff ,
                        0x9a9b9dff9a9c9eff9e9fa1ff9fa0a2ffa1a2a4ffa2a3a5ffa4a5a7ffa4a6a8ff ,
                        0xa7a8aaffaaabadffacadafffaeafb1ffafb0b2ffb0b1b3ffb1b2b4ffb2b3b5ff ,
                        0xb3b4b6ffb5b6b8ffb6b7b9ffb7b8baffb8b9baffb9babcffbbbcbdffbcbdbfff ,
                        0xbdbebfffbdbec0ffbebfc0ffbfc0c1ffc0c1c2ffc1c2c3ffc2c2c4ffc4c4c6ff ,
                        0xc5c6c7ffc6c7c8ffc8c8caffc9c9cbffc9cacbffcacacbffcbcbccffcbcccdff ,
                        0xccccceffcececfffcecfd0ffcfd0d1ffd1d1d2ffd1d2d3ffd2d2d3ffd5d5d6ff ,
                        0xd6d6d7ffd7d7d8ffd7d8d9ffdadbdcffdcddddffdfdfe0ffe0e0e1ffe3e3e4ff ,
                        0xe4e5e5ffe5e5e6ffe5e6e6ffe6e6e7ffe7e7e8ffe8e8e9ffe9e9e9ffeaeaebff ,
                        0xececedffeeeeefffefefeffff0f0f0fff1f1f1fff1f2f2fff2f2f2fff3f3f3ff ,
                        0xf4f4f4fff5f5f5fff5f6f6fff6f6f7fff7f7f7fff9f9f9fffafafafffbfbfbff ,
                        0xfcfcfcfffdfdfdfffefefeffffffffff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff1e0606060606060606060606060606060606060606060606 ,
                        0x06060606060606060606061d3f06060606060606060606060606060606060606 ,
                        0x0606060506060606060606060606063e64060606060606060606060606060606 ,
                        0x0606060604040e120c040406060606060606065f8a1406060606060606060606 ,
                        0x06060606060606053368899385612b040606060606061183a72e060606060606 ,
                        0x060606060606060606060546a5c1c1c1c1c19f3804060606060628a3bc490606 ,
                        0x0606060606060606060606060604219ec1bd825d8bc0c1891305060606063fb9 ,
                        0xc165060606060607090908060605080909091947484a290920a0c1ae2f040606 ,
                        0x06065ac1c1860d06060604247782660e04467e7f81530905437c9a8d558dc1b6 ,
                        0x3903060606087ac1c1a4220606060333b4c1960f42abc1c18e2a0451b6c1c1c1 ,
                        0xc0bac1b63a03060606189ec1c1ba3e0606060332afc19047a5c1c18c270122a2 ,
                        0xc1c1976a8dc0c1b63a0306060636b1c1c1c1580606060332afc1ada7c1c18d28 ,
                        0x03023fbbc1a22d001e96c1b63a030606064ec0c1c1c17b0806060332afc1c1c1 ,
                        0xc1a63301060149bec1870a050476c1b63a030606066fc1c1c1c1a01b06060332 ,
                        0xafc1b0aec1bb6d15040338b5c1aa3c0833a1c1b63a030606108ec1c1c1c1b438 ,
                        0x06060332afc1924ca7c1c07619031385c1c1a884a4c1c1b63a0306062aa7c1c1 ,
                        0xc1c1c15206060332afc1910f40a5c1c1802303359cc1c1c1b7b3c1ba3a030606 ,
                        0x41bbc1c1c1c1c17506060332afc19212043b5a5a5b390604265a7873464b5c56 ,
                        0x1f0506065dc1c1c1c1c1c19817060332afc19212050401010205060604030706 ,
                        0x03030103050606097dc1c1c1c1c1c1ae31060332afc192120506060606060606 ,
                        0x06060606060606060606061aa0c1c1c1c1c1c1bf4c060332afc1921205060606 ,
                        0x06060606060606060606060606060637b2c1c1c1c1c1c1c16c060333b2c19512 ,
                        0x0506060606060606060606060606060606060650c1c1c1c1c1c1c1c18d10051c ,
                        0x59634f0b0606060606060606060606060606060606060670c1c1c1c1c1c1c1c1 ,
                        0xa92a0605030304060606060606060606060606060606060606061290c1c1c1c1 ,
                        0xc1c1c1c1bc440606060606060606060606060606060606060606060606062caa ,
                        0xc1c1c1c1c1c1c1c1c16206060606060606060606060606060606060606060606 ,
                        0x060645bcc1c1c1c1c1c1c1c1c18207060606060606060606060b0b0606060606 ,
                        0x06060606060660c1c1c1c1c1c1c1c1c1c1ab5e555656574d0e06061d5a848860 ,
                        0x2206060a4a575656565699c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c17207061d84 ,
                        0xc1c1c1c18f25060667c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1bd48 ,
                        0x06065dc1c1c1c1c1c16906063db8c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1 ,
                        0xc1c1b236060d89c1c1c1c1c1c19614062cacc1c1c1c1c1c1c1c1c1c1c1c1c1c1 ,
                        0xc1c1c1c1c1c1b134060e8cc1c1c1c1c1c19916062baac1c1c1c1c1c1c1c1c1c1 ,
                        0xc1c1c1c1c1c1c1c1c1c1bb43060664c1c1c1c1c1c17106063bb6c1c1c1c1c1c1 ,
                        0xc1c1c1c1c1c1c1c1c1c1c1c1c1c1c16b06062594c1c1c1c19d2f06065dc1c1c1 ,
                        0xc1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1a42e06062a6e999b7230060622 ,
                        0x9ec1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c17b17060606161706 ,
                        0x06061174c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c0792b06 ,
                        0x06060606062374bec1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1 ,
                        0xc1c19e54270a09225099c1c1c1c1c1c1c1c1c1c1c1c1c1c1
                    End
                    ObjectPalette = Begin
                        0x000301001a15120000000000
                    End
                    ControlTipText ="Weight"
                    Picture ="weight.bmp"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =60
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =720
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =11520
                    Top =60
                    Width =780
                    Height =660
                    TabIndex =5
                    Name ="btnCutSheet"
                    Caption ="Cut Sheet"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000232e0000232e0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffff9fa4f7252fed ,
                        0x262fed262fed262fed262fed262fed262fed262fed262fed262fed262fed262f ,
                        0xed262fed262fed262fed262fed262fed262fed262fed262fed262fed262fed26 ,
                        0x2fed262fed262fed262fed262fed262fed2933eeffffffffffffffffffffffff ,
                        0xffffffffffff9fa3f7242eee242eee242eee242eee242eee2933ee7b81f55e65 ,
                        0xf3232dee242eee242eee242eee757bf47e84f57d83f5666df32d37ef242eee2d ,
                        0x37ef7e84f5525af2242eee242eee242eee242eee242eee242eee242eee2630ed ,
                        0xffffffffffffffffffffffffffffffffffff9fa3f7242eee242eee242eee242e ,
                        0xee242eee2f39eff7f7feb1b4f9232dee242eee242eee232deee7e8fde3e4fd8a ,
                        0x8ff6c9ccfbeff0fe4a53f13942f0ffffff9499f7242eee242eee242eee242eee ,
                        0x242eee242eee242eee2630ed9b9b9b888888888888888888888888888888595e ,
                        0xb9242eee242eee242eee242eee242eee2f39eff7f7fecfd1fb777df46068f328 ,
                        0x32ee232deee7e8fdcbcdfb222cee323beff8f8feb6b9f93841f0ffffffb5b8f9 ,
                        0x676df3545cf2232dee242eee242eee242eee242eee2630ed5d5d5d7d7d7db7b7 ,
                        0xb7b7b7b7b7b7b7b7b7b76e73d2242eee242eee242eee242eee242eee2f39eff7 ,
                        0xf7fed4d6fc898ef6dfe0fcd9dbfc242eeee7e8fdcbcdfb222cee2731eef3f3fe ,
                        0xc8cbfb3740f0ffffffd1d3fca0a5f87d83f5232dee242eee242eee242eee242e ,
                        0xee2630ed5d5d5d959595ffffffffffffffffffffffff9095f7242eee242eee24 ,
                        0x2eee242eee242eee2f39eff7f7feb0b4f9212bee9296f6fefeff343defe7e8fd ,
                        0xcacdfb202aee5159f2fcfcff9b9ff73842f0ffffff9398f7222cee222dee242e ,
                        0xee242eee242eee242eee242eee2630ed5d5d5d959595ffffffffffffffffffff ,
                        0xffff9095f7242eee242eee242eee242eee242eee2f39eff1f2feedeefedbddfc ,
                        0xebecfd8d92f6212beee3e4fdf3f3fee0e1fdecedfea6aaf82b35ef3942f0f9f9 ,
                        0xffecedfedee0fddfe0fd6970f3242eee242eee242eee242eee2630ed5d5d5d95 ,
                        0x9595ffffffffffffffffffffffff9095f7242eee242eee242eee242eee242eee ,
                        0x242eee2c35ee2c35ee2c35ee252fee242eee242eee2b35ee2c35ee2b35ee252f ,
                        0xee232dee242eee252fee2c36ee2c35ee2c35ee2c36ee2731ee242eee242eee24 ,
                        0x2eee242eee2630ed5d5d5d959595ffffffffffffffffffffffffcdcffb9ca0f7 ,
                        0x9ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0 ,
                        0xf79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79ca0f79c ,
                        0xa0f79ca0f75e63b64c50a39b9ff69b9ff69ca1f65d5d5d959595ffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffff8b8b8b6e6e6effffffffffffffffff ,
                        0x5d5d5d959595ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8b8b8b ,
                        0x6e6e6effffffffffffffffff5d5d5d959595ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff8b8b8b6e6e6effffffffffffffffff5d5d5d959595ffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff8b8b8b6e6e6effffffffff ,
                        0xffffffff5d5d5d959595fffffffffffff5f5f5b4b4b4a2a2a2b7b7b7f9f9f9d9 ,
                        0xd9d9b2b2b2fdfdfdebebeba8a8a8f3f3f3bfbfbfacacacacacacaaaaaaebebeb ,
                        0xb7b7b7acacacacacaca8a8a8f3f3f3ffffffb2b2b2d7d7d7ffffffffffffffff ,
                        0xff8b8b8b6e6e6effffffffffffffffff5d5d5d959595ffffffffffffa9a9a995 ,
                        0x9595d5d5d5747474a0a0a0b4b4b4696969fbfbfbd7d7d7545454e7e7e7828282 ,
                        0x808080d0d0d0cfcfcff4f4f47373738f8f8fd0d0d0cececef8f8f8ffffff6868 ,
                        0x68b1b1b1ffffffffffffffffff8b8b8b6e6e6effffffffffffffffff5d5d5d95 ,
                        0x9595fffffffffffffcfcfcbfbfbf8c8c8c5d5d5dbdbdbdb4b4b4626262a0a0a0 ,
                        0x919191565656e7e7e7828282747474a8a8a8cdcdcdffffff7373737e7e7ea7a7 ,
                        0xa7d7d7d7ffffffffffff686868b1b1b1ffffffffffffffffff8b8b8b6e6e6eff ,
                        0xffffffffffffffff5d5d5d959595ffffffffffffb0b0b0666666bebebee9e9e9 ,
                        0xffffffb4b4b4666666d6d6d6bbbbbb555555e7e7e7828282828282d7d7d7e8e8 ,
                        0xe8ffffff737373929292d6d6d6edededffffffffffff686868b1b1b1ffffffff ,
                        0xffffffffff8b8b8b6e6e6effffffffffffffffff5d5d5d959595ffffffffffff ,
                        0xc9c9c9636363949494636363dbdbdbb2b2b2666666fbfbfbd6d6d6515151e7e7 ,
                        0xe77f7f7f6a6a6a9090908f8f8ff3f3f37070707171719090908f8f8fd0d0d08c ,
                        0x8c8c5c5c5c747474969696fdfdfdffffff8b8b8b6e6e6effffffffffffffffff ,
                        0x5d5d5d959595fffffffffffffffffff4f4f4e3e3e3f5f5f5fffffff5f5f5ecec ,
                        0xecfefefefafafae9e9e9fcfcfcefefefeaeaeaeaeaeaeaeaeafdfdfdedededea ,
                        0xeaeaeaeaeaeaeaeaf6f6f6eaeaeaeaeaeaeaeaeaecececffffffffffff8b8b8b ,
                        0x6e6e6effffffffffffffffff5d5d5d959595ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff8b8b8b6e6e6effffffffffffffffff5d5d5d959595ffff ,
                        0xfffffffff5f5f5aaaaaa8888888b8b8bbcbcbcfefefecbcbcb909090eeeeeeff ,
                        0xffffffffffffffffdbdbdb929292929292919191909090dadadafefefec7c7c7 ,
                        0x8f8f8f878787aaaaaaf7f7f7ffffffffffffffffff8b8b8b6e6e6effffffffff ,
                        0xffffffff5d5d5d959595ffffffffffffa5a5a5747474d4d4d4bebebe585858c8 ,
                        0xc8c8b0b0b0585858e6e6e6ffffffffffffffffffc9c9c9595959b0b0b0d2d2d2 ,
                        0xd2d2d2f0f0f0bcbcbc595959b3b3b3d6d6d66d6d6d888888ffffffffffffffff ,
                        0xff8b8b8b6e6e6effffffffffffffffff5d5d5d959595fffffffffffffffffff1 ,
                        0xf1f1bdbdbd8787875a5a5ad0d0d0b0b0b05a5a5a8e8e8e9a9a9abebebefcfcfc ,
                        0xc9c9c9595959adadadcececee4e4e4ffffff808080656565fefefeffffffeeee ,
                        0xeef8f8f8ffffffffffffffffff8b8b8b6e6e6effffffffffffffffff5d5d5d95 ,
                        0x9595ffffffffffffd7d7d76060606a6a6a9d9d9dd6d6d6ffffffb0b0b0595959 ,
                        0xc5c5c5c0c0c05b5b5baaaaaac9c9c95a5a5a8d8d8da0a0a0cbcbcbffffff7d7d ,
                        0x7d686868fffffffffffffefefeffffffffffffffffffffffff8b8b8b6e6e6eff ,
                        0xffffffffffffffff5d5d5d959595ffffffffffffb1b1b1595959dededed9d9d9 ,
                        0xa6a6a6f6f6f6b0b0b0585858dfdfdfe2e2e25e5e5e9e9e9ec9c9c9585858c6c6 ,
                        0xc6f1f1f1f1f1f1fcfcfcb0b0b0595959d0d0d0f0f0f07c7c7c9f9f9fffffffff ,
                        0xffffffffff8b8b8b6e6e6effffffffffffffffff5d5d5d959595ffffffffffff ,
                        0xf3f3f39696966c6c6c707070acacacfdfdfdbdbdbd7676767676767a7a7aa0a0 ,
                        0xa0f4f4f4d2d2d2757575767676767676747474e5e5e5fcfcfcb0b0b07474746a ,
                        0x6a6a8b8b8beaeaeaffffffffffffffffff8b8b8b6e6e6effffffffffffffffff ,
                        0x5d5d5d959595ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8b8b8b ,
                        0x6e6e6effffffffffffffffff5d5d5d959595ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfcfcfcdddddd9999995e5e5e7e7e7effffffffffffffffff5d5d5d959595ffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffdfdfd9e9e9e ,
                        0x8f8f8f9696969090907e7e7e6666665a5a5a5c5c5c5f5f5fe4e4e4ffffffffff ,
                        0xffffffff5d5d5d959595ffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffefefe737373666666a9a9a9adadadbababacdcdcd7d7d7d6060 ,
                        0x60d5d5d5ffffffffffffffffffffffff5d5d5d959595ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff7d7d7d6e6e6efdfdfdffff ,
                        0xffefefef7c7c7c5e5e5ed2d2d2ffffffffffffffffffffffffffffff5d5d5d95 ,
                        0x9595ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffcfc ,
                        0xfc6a6a6a7c7c7cffffffcfcfcf676767707070e2e2e2ffffffffffffffffffff ,
                        0xffffffffffffffff5d5d5d959595ffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffdcdcdc5a5a5aa0a0a09595955a5a5a999999f7f7f7ff ,
                        0xffffffffffffffffffffffffffffffffffffffff5d5d5d959595ffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffdfdfd8080805c5c5c5e5e5e6b ,
                        0x6b6bcdcdcdffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x5d5d5d8d8d8de7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7 ,
                        0xe7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e7e4e4e48e ,
                        0x8e8e5b5b5b636363afafaffdfdfdffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff6c6c6c5c5c5c5a5a5a5a5a5a5a5a5a5a5a5a5a5a ,
                        0x5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a5a ,
                        0x5a5a5a5a5a5a5a5a5a5a5a666666aaaaaaf4f4f4ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Spec Sheet"
                    Picture ="SpecIcon.bmp"

                    LayoutCachedLeft =11520
                    LayoutCachedTop =60
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =720
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7920
                    Top =60
                    Width =780
                    Height =660
                    TabIndex =1
                    Name ="btnTensileStrength"
                    Caption ="Tensile"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000002400000024000000010008000000000010050000c40e0000c40e0000 ,
                        0x0001000000010000424041ff434142ff444243ff454344ff474546ff494849ff ,
                        0x4a4849ff4d4c4dff4f4e4fff525051ff5d5b5cff5e5c5dff676567ff6e6d6dff ,
                        0x6f6e6fff706f70ff757374ff757475ff8e8d8dff908f8fff908f90ff919090ff ,
                        0x949394ff959495ff999798ff999898ff9b9a9bffa7a6a7ffa8a7a8ffa9a8a9ff ,
                        0xaaa9a9ffacabacffb3b3b3ffbdbdbdffc6c5c5ffc7c6c7ffccccccffcdccccff ,
                        0xd2d1d1ffe0dfdfffe9e9e9ffebebebffecececffeeeeeeffefefeffff0f0f0ff ,
                        0xf1f1f1fff2f2f2fff3f3f3fff4f4f4fff5f5f5fff7f7f7fff8f8f8fff9f9f9ff ,
                        0xfdfdfdfffefefeffffffffff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x38383838383838382c3538383838383838383838383838383838322f38383838 ,
                        0x38383838383838383838331101082338383838383838383838383838381f0401 ,
                        0x15383838383838383838383838381e0000000938383838383838383838383838 ,
                        0x2901000000253838383838383838383838381a000000022e3838383838383838 ,
                        0x3838383827000000002138383838383838383834242e1a000000022e38383838 ,
                        0x38383838383838382700000000212825373838383838380f000b19000000022e ,
                        0x38383838383838383838383827000000001d0600173838383838380300031400 ,
                        0x0000022e38383838383838383838383827000000001200000d38383838383803 ,
                        0x000314000000022e38383838383838383838383827000000001200000d383838 ,
                        0x38383803000314000000022e3838383838383838383838382700000000120000 ,
                        0x0d38383838383803000314000000022e38383838383838383838383827000000 ,
                        0x001200000d38383838260c00000314000000021c0c0c0c0c0c0c0c0c0c0c0c0c ,
                        0x2000000000120000060c2b383823000000031400000002170000000000000000 ,
                        0x000000001b0000000012000000002a3838230000000314000000021700000000 ,
                        0x00000000000000001b0000000012000000002a3838260c00000314000000021c ,
                        0x0c0c0c0c0c0c0c0c0c0c0c0c2000000000120000060c2b383838380300031400 ,
                        0x0000022e38383838383838383838383827000000001200000d38383838383803 ,
                        0x000314000000022e38383838383838383838383827000000001200000d383838 ,
                        0x38383803000314000000022e3838383838383838383838382700000000120000 ,
                        0x0d38383838383803000314000000022e38383838383838383838383827000000 ,
                        0x001200000d3838383838380e000a18000000022e383838383838383838383838 ,
                        0x27000000001d05001638383838383833242e1a000000022e3838383838383838 ,
                        0x383838382700000000212825363838383838383838381a000000013038383838 ,
                        0x38383838383838382700000000213838383838383838383838381d0000000938 ,
                        0x3838383838383838383838382901000000243838383838383838383838383210 ,
                        0x00072238383838383838383838383838381e0301133738383838383838383838 ,
                        0x38383838293538383838383838383838383838383838312d3838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x3838383838383838383838383838383838383838383838383838383838383838 ,
                        0x383838383838383838383838383838383838383838383838
                    End
                    ObjectPalette = Begin
                        0x000301004140420000000000
                    End
                    ControlTipText ="Tensile Strength"
                    Picture ="StrengthIcon.bmp"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =60
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =720
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8820
                    Top =60
                    Width =780
                    Height =660
                    TabIndex =2
                    Name ="btnBOMExplode"
                    Caption ="Bom Explode"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000c40e0000c40e0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffff7f7f7c7c7c7a7a7 ,
                        0xa79f9f9faeaeaed4d4d4fbfbfbffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffbfbfba8a8 ,
                        0xa84343430a0a0a020202020202020202020202020202111111595959c4c4c4ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffc4c4c42d2d2d02020202020202020202020202020202020202020202 ,
                        0x02020202020202020202024c4c4ce3e3e3ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff90909005050502020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202131313bdbdbd ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff89898902020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020a0a0abdbdbdffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffb1 ,
                        0xb1b1020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202020202020202020202020202020202020202121212dfdfdfffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffefefef171717020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202474747fdfdfdffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff7f7f7f020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020202020202020202020202bdbdbdffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffdfdfd ,
                        0x1818180202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202515151ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffc1c1c10202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020202020b0b0bf7f7f7ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff8787870202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202c2c2c2 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xff62626202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202020202a3a3a3ffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff5c5c5c02020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202020202020202020202020202020202919191ffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffff5d5d5d02020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x029c9c9cffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffff777777020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020202b3b3b3ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffaaaaaa020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020202020202020202020202030303e7e7e7ffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffededed ,
                        0x0505050202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202313131ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff5555550202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202929292ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffcfcfcf0404040202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020202020202020202020202021d1d1df4f4f4ffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffff70707002020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x030303adadadffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffff6f6f63c3c3c02020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202020202020202727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffededed3a3a3a020202020202020202020202020202020202020202020202 ,
                        0x020202020202020202020202020202020202020202696969fdfdfdffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffff7f7f76c6c6c030303020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020d0d0d9898 ,
                        0x98fdfdfdffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffc9c9c94b4b4b0202020202020202020202020202020202020202020202 ,
                        0x02080808676767e2e2e2ffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffe2e2e29f9f9f6a6a6a4c4c ,
                        0x4c444444535353757575acacacf1f1f1ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffa1a1a19b9b9bb8b8b8b8b8b8b3b3b39292927f7f7fffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffff52525200000000000000000000000000 ,
                        0x00002f2f2fffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff52525200 ,
                        0x00000000000000000000000000002f2f2fffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffd8d8d8ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffa1a1a17676767676762121217676767676768e8e8effffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffe1e1e19e9e9e525252afaf ,
                        0xafa6a6a6ffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff515151 ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffff8f8f8bdbdbd8c8c ,
                        0x8c7777770f0f0f000000000000a5a5a5ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff828282bfbfbfffffffffffffffffffffffffffffffd1d1 ,
                        0xd1646464565656909090bdbdbd6e6e6e0f0f0f0000000000003e3e3ecececeff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffebebeb434343e6e6e6ffff ,
                        0xfffefefecccccc5858586c6c6cddddddfffffffffffffffffff8f8f829292911 ,
                        0x11111e1e1e949494ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffdbdbdb5e5e5e4949494f4f4f7c7c7cdededeffffffffffffffffffff ,
                        0xffffffffffebebebededed7e7e7ef0f0f0d8d8d8ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffff6f6f6ffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Bom Explode"
                    Picture ="Bom.bmp"

                    LayoutCachedLeft =8820
                    LayoutCachedTop =60
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =720
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9000
                    Top =780
                    Width =480
                    TabIndex =11
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

                    LayoutCachedLeft =9000
                    LayoutCachedTop =780
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =20520
                    Top =1680
                    Width =720
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label73"
                    Caption ="OD"
                    LayoutCachedLeft =20520
                    LayoutCachedTop =1680
                    LayoutCachedWidth =21240
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
                    TabIndex =13
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbOracleStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblDesignCodeJacket.JacketDescription_SalesForce, tblDesignTypes"
                        ".SalesForceDescription FROM tblDesignCodeJacket INNER JOIN (tblCableConstruction"
                        "References INNER JOIN tblDesignTypes ON tblCableConstructionReferences.DesignTyp"
                        "eID = tblDesignTypes.DesignTypeID) ON (tblDesignCodeJacket.JacketLetter = tblCab"
                        "leConstructionReferences.CableLevel1) AND (tblCableConstructionReferences.CableF"
                        "amily = tblDesignCodeJacket.CableLetter) WHERE (((tblDesignTypes.SalesForceDescr"
                        "iption)=[Forms]![frmFindConstruction]![cbFamilyDesc])); "

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
                    TabIndex =14
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
                    Left =10620
                    Top =60
                    Width =780
                    Height =660
                    TabIndex =4
                    Name ="btnWhereUsed"
                    Caption ="Where Used"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000130b0000130b0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffa5a5a50a0a0a343434e9e9e9ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffa4a4a4030303000000000000454545 ,
                        0xfffffffffffffafafacbcbcbbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbd ,
                        0xbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbddfdfdfff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffa4a4a4030303 ,
                        0x000000000000000000363636ffffffa6a6a61616160000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000181818a9a9a9fefefeffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffa4a4a4030303000000000000000000181818d6d6d6b4b4b40101010000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000373737e9e9e9fffffffafafaa4a4a45454542b2b2b ,
                        0x2626263d3d3d7a7a7adadadaf2f2f2080808000000000000000000181818d6d6 ,
                        0xd6ffffff3636360000000000003b3b3b72727272727272727272727272727272 ,
                        0x7272727272727272727272727272727272727272787878f2f2f2ffffffc5c5c5 ,
                        0x2525250000000000000000000000000000000000000404046d6d6d0505050000 ,
                        0x00000000181818d6d6d6ffffffffffff0202020000001d1d1dfbfbfbffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffb4b4b40505050000000f0f0f6f6f6fadadadb9b9b99292923636 ,
                        0x36000000000000000000131313282828d6d6d6ffffffffffffffffff00000000 ,
                        0x0000303030ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffdcdcdc0c0c0c000000434343e9e9e9ffff ,
                        0xfffffffffffffffffffffefefe9e9e9e060606000000737373ffffffffffffff ,
                        0xffffffffffffffff000000000000303030ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffff5757570000 ,
                        0x00303030f6f6f6ffffffffffffffffffffffffffffffffffffffffff9e9e9e00 ,
                        0x0000040404dadadaffffffffffffffffffffffff000000000000303030ffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffebebeb030303000000c3c3c3ffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffefefe3636360000007a7a7affffffffffffffffffffffff ,
                        0x000000000000303030ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffb2b2b20000001b1b1bfefefeffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffff9292920000003d3d3d ,
                        0xffffffffffffffffffffffff000000000000303030ffffffffffffffffffe4e4 ,
                        0xe40a0a0a0a0a0a535353ffffffffffff7777770a0a0a777777ffffff97979700 ,
                        0x0000424242ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffb9b9b9000000262626ffffffffffffffffffffffff0000000000003030 ,
                        0x30ffffffffffffffffffe3e3e30000000000004c4c4cffffffffffff72727200 ,
                        0x00005e5e5effffff9f9f9f000000343434ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffadadad0000002b2b2bffffffffffffffff ,
                        0xffffffff000000000000303030ffffffffffffffffffe7e7e726262626262666 ,
                        0x6666ffffffffffff878787262626565656ffffffc9c9c9000000070707f1f1f1 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffff6f6f6f0000 ,
                        0x00545454ffffffffffffffffffffffff000000000000303030ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfcfcfc1c1c1c000000848484ffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffe9e9e90f0f0f000000a4a4a4ffffffffffffffffffffffff00000000 ,
                        0x0000303030ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffff949494000000060606bababaffffffffff ,
                        0xfffffffffffffffffffffffffff6f6f6434343000000252525fafafafefefedf ,
                        0xdfdfffffffffffff000000000000303030fffffffffffffffffff4f4f4979797 ,
                        0x979797b6b6b6ffffffffffffc6c6c69797979797979f9f9ffffffffdfdfd4545 ,
                        0x45000000060606848484f1f1f1fffffffffffffefefec3c3c330303000000005 ,
                        0x0505c5c5c5ffffffaaaaaabdbdbdffffffffffff000000000000303030ffffff ,
                        0xffffffffffffe3e3e30000000000004c4c4cffffffffffff7272720000000000 ,
                        0x000000009f9f9ffffffff0f0f04545450000000000000707073434344242421b ,
                        0x1b1b0000000000000c0c0cb4b4b4ffffffeeeeee1f1f1fbdbdbdffffffffffff ,
                        0x000000000000303030ffffffffffffffffffe3e3e30000000000004c4c4cffff ,
                        0xffffffff7272720000000000000000000a0a0ac1c1c1fffffffdfdfd9494941c ,
                        0x1c1c000000000000000000000000030303575757dcdcdcfffffff4f4f4424242 ,
                        0x000000bdbdbdffffffffffff000000000000303030fffffffffffffffffff4f4 ,
                        0xf4979797979797b6b6b6ffffffffffffc6c6c69797979797979797979797979e ,
                        0x9e9efafafafffffffffffffcfcfcc9c9c99f9f9f979797b2b2b2ebebebffffff ,
                        0xffffffffffff787878000000000000bdbdbdffffffffffff0000000000003030 ,
                        0x30ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffff727272000000000000bdbdbdffff ,
                        0xffffffff000000000000303030ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7272 ,
                        0x72000000000000bdbdbdffffffffffff000000000000303030ffffffffffffff ,
                        0xffffe7e7e7262626262626666666ffffffffffff878787262626262626262626 ,
                        0x262626262626262626262626262626262626262626262626262626262626afaf ,
                        0xafffffffffffffffffff727272000000000000bdbdbdffffffffffff00000000 ,
                        0x0000303030ffffffffffffffffffe3e3e30000000000004c4c4cffffffffffff ,
                        0x7272720000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000a1a1a1ffffffffffffffffff727272000000000000bd ,
                        0xbdbdffffffffffff000000000000303030ffffffffffffffffffe4e4e4060606 ,
                        0x060606505050ffffffffffff7575750606060606060606060606060606060606 ,
                        0x06060606060606060606060606060606060606060606a3a3a3ffffffffffffff ,
                        0xffff727272000000000000bdbdbdffffffffffff000000000000303030ffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffff727272000000000000bdbdbdffffffffffff ,
                        0x000000000000303030ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffff727272000000 ,
                        0x000000bdbdbdffffffffffff000000000000303030ffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff727272000000000000bdbdbdffffffffffff0000000000003030 ,
                        0x30ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffff727272000000000000bdbdbdffff ,
                        0xffffffff000000000000303030ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7272 ,
                        0x72000000000000bdbdbdffffffffffff00000000000015151572727272727272 ,
                        0x7272727272727272727272727272727272727272727272727272727272727272 ,
                        0x7272727272727272727272727272727272727272727272727272727272727272 ,
                        0x72727272727272727272333333000000000000bdbdbdffffffffffff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000000000bd ,
                        0xbdbdffffffffffff070707000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000cbcbcbffffffffffff4b4b4b000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000161616fafafaffffffffffff ,
                        0xd4d4d41010100000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000010101 ,
                        0xa6a6a6ffffffffffffffffffffffffd4d4d44b4b4b0707070000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000020202363636b4b4b4ffffffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Where Used"
                    Picture ="WhereUsed.bmp"

                    LayoutCachedLeft =10620
                    LayoutCachedTop =60
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =720
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9720
                    Top =60
                    Width =780
                    Height =660
                    TabIndex =3
                    Name ="btnBomCostExplode"
                    Caption ="Bom Cost Explode"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000002400000024000000010008000000000010050000130b0000130b0000 ,
                        0x0001000000010000000000ff010101ff020202ff030303ff040404ff050505ff ,
                        0x060606ff070707ff080808ff090909ff0a0a0aff0b0b0bff0c0c0cff0d0d0dff ,
                        0x0e0e0eff0f0f0fff101010ff111111ff121212ff131313ff141414ff151515ff ,
                        0x161616ff171717ff181818ff191919ff1a1a1aff1b1b1bff1c1c1cff1d1d1dff ,
                        0x1e1e1eff1f1f1fff202020ff212121ff222222ff232323ff242424ff252525ff ,
                        0x262626ff272727ff282828ff292929ff2a2a2aff2b2b2bff2c2c2cff2e2e2eff ,
                        0x2f2f2fff303030ff313131ff323232ff333333ff343434ff363636ff373737ff ,
                        0x393939ff3a3a3aff3b3b3bff3c3c3cff3d3d3dff3f3f3fff404040ff414141ff ,
                        0x424242ff434343ff444444ff454545ff464646ff474747ff484848ff494949ff ,
                        0x4a4a4aff4b4b4bff4d4d4dff4e4e4eff4f4f4fff515151ff525252ff535353ff ,
                        0x555555ff565656ff575757ff595959ff5a5a5aff5c5c5cff5d5d5dff5e5e5eff ,
                        0x5f5f5fff606060ff616161ff626262ff636363ff646464ff656565ff666666ff ,
                        0x676767ff686868ff6a6a6aff6c6c6cff6e6e6eff6f6f6fff707070ff717171ff ,
                        0x737373ff747474ff757575ff767676ff777777ff787878ff797979ff7a7a7aff ,
                        0x7b7b7bff7c7c7cff7d7d7dff7e7e7eff7f7f7fff808080ff818181ff828282ff ,
                        0x838383ff848484ff858585ff868686ff878787ff8a8a8aff8b8b8bff8c8c8cff ,
                        0x8d8d8dff909090ff919191ff929292ff949494ff959595ff969696ff979797ff ,
                        0x989898ff999999ff9a9a9aff9b9b9bff9c9c9cff9d9d9dff9e9e9eff9f9f9fff ,
                        0xa0a0a0ffa1a1a1ffa2a2a2ffa5a5a5ffa6a6a6ffa7a7a7ffa8a8a8ffa9a9a9ff ,
                        0xaaaaaaffacacacffadadadffaeaeaeffb1b1b1ffb2b2b2ffb3b3b3ffb4b4b4ff ,
                        0xb5b5b5ffb6b6b6ffb7b7b7ffb8b8b8ffbababaffbcbcbcffbdbdbdffbebebeff ,
                        0xc0c0c0ffc2c2c2ffc3c3c3ffc4c4c4ffc6c6c6ffc7c7c7ffc8c8c8ffc9c9c9ff ,
                        0xcbcbcbffccccccffcdcdcdffcececeffcfcfcfffd0d0d0ffd1d1d1ffd2d2d2ff ,
                        0xd3d3d3ffd4d4d4ffd5d5d5ffd6d6d6ffd7d7d7ffd8d8d8ffdadadaffdbdbdbff ,
                        0xdcdcdcffddddddffdededeffdfdfdfffe0e0e0ffe1e1e1ffe2e2e2ffe3e3e3ff ,
                        0xe4e4e4ffe5e5e5ffe6e6e6ffe7e7e7ffe8e8e8ffe9e9e9ffeaeaeaffebebebff ,
                        0xecececffedededffefefeffff0f0f0fff1f1f1fff2f2f2fff3f3f3fff4f4f4ff ,
                        0xf5f5f5fff6f6f6fff7f7f7fff8f8f8fff9f9f9fffafafafffbfbfbfffcfcfcff ,
                        0xfdfdfdfffefefeffffffffff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ff000000ff000000ff000000ff000000ff000000ff000000ff ,
                        0x000000ff000000ffe0e0e0e0e0e0e0e0e0e0dfdddbdadbdddfe0e0e0e0e0e0e0 ,
                        0xe0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0d5bda4856b64708aaac3dae0 ,
                        0xe0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0d6a065291329393d3625 ,
                        0x113874a8dce0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0b449142280aec0 ,
                        0xc9cbc7bea96716195fcce0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0b72b0c6d ,
                        0xd7e0e0e0e0e0e0e0e0e0c754074ec8e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0c5 ,
                        0x36279edee0e0c5b5b0b4c0dee0e0e0de80145acce0e0e0e0e0e0e0e0e0e0e0e0 ,
                        0xe0e0c6451fa0dce0d08b3f12010d317bb2e0e0e0d87d125fd7e0e0e0e0e0e0e0 ,
                        0xe0e0e0e0e0db570592dde0782e08113daf6f17031f92e0e0e0d666087fe0e0e0 ,
                        0xe0e0e0e0e0e0e0e0e0a41b4ddfe057041e99e0d9d3dce0e031043bd6e0e0d52b ,
                        0x2bc9e0e0e0e0e0e0e0e0e0e0dd6b11b8daba2936e0e0e06e0f7fd8e0cb4d07a8 ,
                        0xe0e0e0961198e0e0e0e0e0e0e0e0e0e0cd414fd9c12b2ad5e0b5811a02186ba3 ,
                        0xe0b93508b2e0e0d8256dd8e0e0e0e0e0e0e0e0e0bb1d96deb20c9edecd4b0d24 ,
                        0x0836141ee0e0bf0a47e0e0e06445c9e0e0e0e0e0e0e0e0e0ac0fafdf3f14b6e0 ,
                        0xb81b5c7512956a0472e0e02014a1e0e08d2cbee0e0e0e0e0e0e0e0e0a517b7cf ,
                        0x173ec6e0ce8fb66a0c76410491e0e0670687e0e09c1fbae0e0e0e0e0e0e0e0e0 ,
                        0xa518b7a3045ee0e0d79a471501090d45e0e0e0e0017ee0e09e1ebae0e0e0e0e0 ,
                        0xe0e0e0e0ab10afb9113cc8e0c43211280772b4cce0e0e06f0584e0e08f2bbde0 ,
                        0xe0e0e0e0e0e0e0e0ba1a9ad93616b4e0b8154d7912915931e0e0e02413a0e0e0 ,
                        0x6941c7e0e0e0e0e0e0e0e0e0c83861da460a93decb4220510c431518e0e0c20a ,
                        0x3fe0e0dd3562d4e0e0e0e0e0e0e0e0e0d96020c38f2038d4e0a1330500085690 ,
                        0xe0e04d059ce0e0aa0f8edfe0e0e0e0e0e0e0e0e0e09f1958ce420d2ee0e0c365 ,
                        0x0f7cd6e0e0820c46e0e0d63628c3e0e0e0e0e0e0e0e0e0e0e0db4b039db82b05 ,
                        0x249dd4d7e0dbd4d34d0937d6e0dd700677e0e0e0e0e0e0e0e0e0e0e0e0e0b623 ,
                        0x2dbaaa3a140b1d4ba675250a1085e0e0dd9b1b35d3e0e0e0e0e0e0e0e0e0e0e0 ,
                        0xe0e0e0a2264ab2c55d1b1c0d020a226488bfe0df9c2d3ab6e0e0e0e0e0e0e0e0 ,
                        0xe0e0e0e0e0e0e0dfa72a2179d2bf8c5d4055addee0e0c35e1537b5dfe0e0e0e0 ,
                        0xdfbed2e0e0e0e0e0e0e0e0e0e0b134062490cdd7d9dad8d6bf681106162abce0 ,
                        0xe0dfd9ddcd416fdbdddbe0e0e0e0e0e0e0e0d07b2c12163c535c4c300a153383 ,
                        0xab1593dee0c95b94c43d67d28685d8e0e0e0e0e0e0e0e0e0c78f663e042b4450 ,
                        0x6998cee0c4353dd7e0c74c0f8a719480137cd7e0e0e0e0e0e0e0e0e0e0e0ddbf ,
                        0x097acdd4dee0da9a3f0e56d9e0e0ca75a3d5db9881d3e0e0e0e0e0e0e0e0e0e0 ,
                        0xe0e0e0e04156c8d3ab681e051a89d8e0e0e0e0dbc468a3de85415bd7e0e0e0e0 ,
                        0xe0e0e0e0e0e0e0e0820d1d0d0c07106bc0e0e0e0dfcfb3721b1b89dc642b38d2 ,
                        0xe0e0e0e0e0e0e0e0e0e0e0e0c165636e814a2197d2ded3b18b5218255da2d89e ,
                        0x8bc2d2dfe0e0e0e0e0e0e0e0e0e0e0e0e0dddedfe0b733122e3a2f1d0f4292bd ,
                        0xce759a6c1086dae0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0d27f3b223c74 ,
                        0xb9dee0e0ca3c66d95a60d1e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0 ,
                        0xe0e0e0e0e0e0e0e0cd4873dbdbd6e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0 ,
                        0xe0e0e0e0e0e0e0e0e0e0e0e0dba2b9dfe0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0 ,
                        0xe0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0e0
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="BOM Cost Explode"
                    Picture ="BomCost.bmp"

                    LayoutCachedLeft =9720
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =720
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =27000
                    Top =120
                    Width =720
                    Height =660
                    TabIndex =6
                    Name ="btnFiberClose"
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

                    LayoutCachedLeft =27000
                    LayoutCachedTop =120
                    LayoutCachedWidth =27720
                    LayoutCachedHeight =780
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
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =26460
                    Top =60
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =17
                    Name ="txItemID"
                    ControlSource ="ItemID"
                    Format ="General Number"

                    LayoutCachedLeft =26460
                    LayoutCachedTop =60
                    LayoutCachedWidth =26820
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =24360
                    Top =60
                    Width =1260
                    Height =300
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =500
                    TabIndex =14
                    Name ="txColor"
                    ControlSource ="Color"

                    LayoutCachedLeft =24360
                    LayoutCachedTop =60
                    LayoutCachedWidth =25620
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22440
                    Top =60
                    Width =1860
                    Height =300
                    ColumnWidth =1650
                    ColumnOrder =4
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="txBase"
                    ControlSource ="Base"

                    LayoutCachedLeft =22440
                    LayoutCachedTop =60
                    LayoutCachedWidth =24300
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
                    Left =16500
                    Top =60
                    Width =1380
                    Height =300
                    ColumnOrder =5
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="txFiber1"
                    ControlSource ="Fiber"

                    LayoutCachedLeft =16500
                    LayoutCachedTop =60
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10260
                    Top =60
                    Width =2940
                    Height =300
                    ColumnWidth =2835
                    ColumnOrder =6
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="txCustomer"
                    ControlSource ="Customer"

                    LayoutCachedLeft =10260
                    LayoutCachedTop =60
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25680
                    Top =60
                    Width =720
                    Height =300
                    ColumnOrder =7
                    FontSize =9
                    FontWeight =500
                    TabIndex =15
                    Name ="txUOM"
                    ControlSource ="PrintSpacing"

                    LayoutCachedLeft =25680
                    LayoutCachedTop =60
                    LayoutCachedWidth =26400
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
                    TabIndex =1
                    Name ="txItemNo"
                    ControlSource ="ItemNo"
                    OnDblClick ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7140
                    Width =426
                    Height =426
                    TabIndex =19
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

                    LayoutCachedLeft =7140
                    LayoutCachedWidth =7566
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
                    TabIndex =3
                    Name ="txActive"
                    ControlSource ="Active"
                    Format ="Yes/No"

                    LayoutCachedLeft =7620
                    LayoutCachedTop =60
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =26460
                    Top =60
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =16
                    Name ="txLocation"
                    ControlSource ="Location"

                    LayoutCachedLeft =26460
                    LayoutCachedTop =60
                    LayoutCachedWidth =26820
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17940
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="txFiber2"
                    ControlSource ="Fiber2"

                    LayoutCachedLeft =17940
                    LayoutCachedTop =60
                    LayoutCachedWidth =19200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13260
                    Top =60
                    Width =2220
                    Height =300
                    ColumnWidth =2130
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="txPID"
                    ControlSource ="PID"

                    LayoutCachedLeft =13260
                    LayoutCachedTop =60
                    LayoutCachedWidth =15480
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15540
                    Top =60
                    Width =900
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="txCustomerRevision"
                    ControlSource ="CustomerRev"

                    LayoutCachedLeft =15540
                    LayoutCachedTop =60
                    LayoutCachedWidth =16440
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19260
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="txJacket"
                    ControlSource ="JacketMaterial"

                    LayoutCachedLeft =19260
                    LayoutCachedTop =60
                    LayoutCachedWidth =20460
                    LayoutCachedHeight =360
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
                    TabIndex =4
                    Name ="txReason"
                    ControlSource ="Reason"

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
                    Left =5880
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txCreationDate"
                    ControlSource ="CreationDate"
                    Format ="Medium Date"
                    ShowDatePicker =0

                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21300
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =12
                    Name ="txSetup"
                    ControlSource ="StandardOperation"

                    LayoutCachedLeft =21300
                    LayoutCachedTop =60
                    LayoutCachedWidth =22380
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20520
                    Top =60
                    Width =720
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="Text74"
                    ControlSource ="NominalOD"

                    LayoutCachedLeft =20520
                    LayoutCachedTop =60
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =26880
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =18
                    Name ="txOracleStatus"
                    ControlSource ="OracleStatus"

                    LayoutCachedLeft =26880
                    LayoutCachedTop =60
                    LayoutCachedWidth =27960
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =420
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =60
                    Width =4380
                    Height =240
                    BorderColor =-2147483633
                    ForeColor =4144959
                    Name ="lbMessage"
                    Caption ="| Type to search  | Enter to select  | Esc to cancel"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =300
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
Forms.frmBOMCostExplode.Refresh
End Sub

Private Sub btnBOMExplode_Click()
DoCmd.SetWarnings False
Call BomComponentExplode(Me.Oracle)

DoCmd.OpenForm "frmBOMComponentExplode"

Forms!frmBOMComponentExplode!lblItem.Caption = Me.Oracle

End Sub

Private Sub btnCutSheet_Click()
cutSheet.OpenCutSheet (Me.Oracle)
'OpenCutSheet (Me.Oracle)
End Sub

Private Sub btnExcelExport_Click()


On Error GoTo ErrorHandler

DoCmd.OutputTo acOutputForm, "frmFindItem", acFormatXLS, Environ("USERPROFILE") & "\Desktop\ItemSearchResults" & Format(Now, "yyyyMMddhhmmss") & ".xls"
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

On Error GoTo ErrorHandler
If CloseClient(Me.hWnd) Then
    Exit Sub
End If

DoCmd.Close acForm, Me.name

ErrorExit:
Exit Sub

ErrorHandler:
DoCmd.Close acForm, Me.name
Resume ErrorExit

End Sub

Private Sub btnFind_Click()
Call OpenAClient(1, Me.txItemID, Nz(Me.Location, 2))
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

On Error GoTo ErrorHandler:

DoCmd.Echo False




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
strFilter = strFilter & " OR [DesignCode] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [PID] like """ & Me.ItemNoSearch & "*"")"

'If Me.tglShowObsoleteItems <> 0 Then
'Else
'strFilter = strFilter & " AND [OracleStatus] = ""active"""
'End If



Me.Filter = strFilter
Me.FilterOn = True
'Call Refresh_Click
DoCmd.Echo True

ErrorExit:
Exit Sub

ErrorHandler:
DoCmd.Echo True
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit
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
Call UserCheck
End Sub

Private Sub Form_Load()
Call UserCheck
DoCmd.SetWarnings False
'Call btnSearch_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
'Call btnSearch_Click
Call btnRemoveFilter_Click
If sUser = 1 Then
btnCutSheet.Visible = True
Else
btnCutSheet.Visible = False
End If

Me.ItemNoSearch.SetFocus
btnSortDescending_Click
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

Private Sub txItemNo_DblClick(Cancel As Integer)
Call btnFind_Click
End Sub

Private Sub txItemNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call btnBomCostExplode_Click
End If
End Sub

Private Sub txOracle_DblClick(Cancel As Integer)
Call btnFind_Click
End Sub

Private Sub Weight_Click()


Dim Weight As Double

Weight = BOM_WeightCalculator(Me.Oracle)

If Weight = 0 Then
    Exit Sub
End If

MsgBox Me.Oracle & " weighs " & vbCrLf & Weight & " kg/km" & vbCrLf & RoundRanges(Weight * 0.671969) & " lb/kft"


End Sub

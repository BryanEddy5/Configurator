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
    Width =21120
    DatasheetFontHeight =10
    ItemSuffix =88
    Left =-26955
    Top =5595
    Right =-5835
    Bottom =12420
    DatasheetGridlinesColor =12632256
    Filter ="[FamilyID] like \"**\" OR [FamilyDescription] like   \"**\" OR [NewProductCode] "
        "like  \"**\" OR [EngineeringChange] like   \"**\" OR [NewProductRelease] like   "
        "\"**\""
    OrderBy ="[tblCableFamily].[FamilyID], [tblCableFamily].[FamilyDescription]"
    RecSrcDt = Begin
        0x948295fd8fcde440
    End
    RecordSource ="tblCableFamily"
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
                    Left =240
                    Top =1680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="FamilyID"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =6540
                    Top =1680
                    Width =4260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="New Product Code"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =12900
                    Top =1680
                    Width =2220
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label24"
                    Caption ="New Product Release"
                    LayoutCachedLeft =12900
                    LayoutCachedTop =1680
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =180
                    Top =420
                    Width =2640
                    Name ="Box35"
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =1140
                End
                Begin Label
                    OverlapFlags =85
                    Left =10920
                    Top =1680
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Engineering Change"
                    LayoutCachedLeft =10920
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1260
                    Width =15240
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =15420
                    LayoutCachedHeight =1260
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8340
                    Top =780
                    Width =1320
                    TabIndex =3
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =780
                    LayoutCachedWidth =9660
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
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Cable Family"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =300
                    LayoutCachedTop =540
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =1035
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =6840
                    Top =780
                    Width =1320
                    Name ="btnSearch"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =780
                    LayoutCachedWidth =8160
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
                    Left =3000
                    Top =780
                    Width =3600
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =780
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16440
                    Top =1620
                    Width =480
                    TabIndex =5
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

                    LayoutCachedLeft =16440
                    LayoutCachedTop =1620
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =1980
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =3000
                    Top =60
                    Width =3900
                    Height =600
                    FontSize =9
                    Name ="Label65"
                    Caption ="Master Search: Searches all fields displayed with embedded wild card."
                    LayoutCachedLeft =3000
                    LayoutCachedTop =60
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =660
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9840
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

                    LayoutCachedLeft =9840
                    LayoutCachedTop =780
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =1080
                    Top =1680
                    Width =5400
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label68"
                    Caption ="Family Description"
                    LayoutCachedLeft =1080
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =15180
                    Top =1680
                    Width =1215
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label85"
                    Caption ="Date Created"
                    LayoutCachedLeft =15180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =16395
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10440
                    Top =660
                    Width =960
                    Height =480
                    TabIndex =1
                    Name ="btnAddDocs"
                    Caption ="Add New NPR"
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
                        0xfffffffffffffffffffffffffffffffffffffffffff3f3f3cccdcdc9cacac9ca ,
                        0xcac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9 ,
                        0xcacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9caca ,
                        0xcacacae7e7e7ffffffffffffffffffffffffffffffffffffffffffffffffc3c3 ,
                        0xc35a5b5b56575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x565757565757565757565757565757575858adadadffffffffffffffffffffff ,
                        0xfffffffffffffff0f1f15b5c5c56575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57d2d2d2ffffffffffffffffffffffffffffffc4c4c456575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x575657575657575657575657577f8080ffffffffffffffffffffffffffffffb1 ,
                        0xb1b1565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757565757565757565757565757565757565757f4f4f4ff ,
                        0xffffffffffffffffffffffabacac565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757d8d8d8ffffffffffffffffffffffffabacac565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757babbbbffffffffffffffffff ,
                        0xffffffabacac5657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x9c9c9cffffffffffffffffffffffffabacac5657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657577f8080ffffffffffffffffffffffffabacac5657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x565757565757565757565757565757565757565757565757636363ffffffffff ,
                        0xffffffffffffffabacac56575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757f1f1f1ffffffffffffffffffabacac56575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757565757565757565757d1d1d1ffffffffffffffffffab ,
                        0xacac565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757565757565757565757565757565757565757565757b2 ,
                        0xb3b3ffffffffffffffffffabacac565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757969696ffffffffffffffffffabacac565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757777878ffffffffffff ,
                        0xffffffabacac5657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575e5f5ffdfdfdffffffffffffabacac5657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x565757565757565757565757565757565757eaeaeaffffffffffffabacac5657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x565757565757565757565757565757565757565757565757565757565757c9c9 ,
                        0xc9ffffffffffffabacac56575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757b0b1b1ffffffffffffabacac5657578d8e8e595a5a56 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757565757565757565757565757bababaffffffffffffaa ,
                        0xabab565757b6b7b7d5d5d5727373555656565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575770 ,
                        0x7070f4f4f4ffffffffffffa9a9a9565757b6b7b7ffffffffffffe7e7e7dbdbdb ,
                        0xdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdb ,
                        0xdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdb ,
                        0xdbdbdbdbdbd7d7d7cccdcdfcfcfcffffffffffffffffffa9a9a9565757afb0b0 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffb4b4b4e1e1e1ffffffffffffffffff ,
                        0xffffffa9a9a95657576b6b6beaebebffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffff8f8f8afafaf585959 ,
                        0xe1e1e1ffffffffffffffffffffffffb4b4b4565757565757595a5a5f60605f60 ,
                        0x605f60605f60605f60605f60605f60605f60605f60605f60605f60605f60605f ,
                        0x60605f60605f6060cfcfcfffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffa9aaaa565757565757e5e5e5ffffffffffffffffffffffffdfdfdf5657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x57575657575657575657575657575657575657576c6d6dc5c5c5d1d1d1d1d1d1 ,
                        0xd1d1d1d1d1d1d1d1d1d1d1d1b8b9b95d5e5e565757696a6affffffffffffffff ,
                        0xffffffffffffffffffffa2a3a356575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575b5c ,
                        0x5cd4d4d4ffffffffffffffffffffffffffffffffffffffffffcecece9d9d9d91 ,
                        0x9292919292919292919292919292919292919292919292919292919292919292 ,
                        0x9192929192929192929192926465655657575657575657575657575657575657 ,
                        0x575657575657577c7d7de5e6e6ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffaaabab5657575657 ,
                        0x57565757565757565757565757565757565757cecfcfffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffdfdfdd7d7d7c9cacac9cacac9cacac9cacac9cacac9cacae1e1e1ff ,
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
                    ControlTipText ="Add NPR Documents"
                    Picture ="NPR.bmp"

                    LayoutCachedLeft =10440
                    LayoutCachedTop =660
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =17400
                    Top =1500
                    Width =600
                    Height =480
                    FontSize =10
                    FontWeight =500
                    Name ="Label87"
                    Caption ="Open Bases"
                    LayoutCachedLeft =17400
                    LayoutCachedTop =1500
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =1980
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
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =780
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="Text6"
                    ControlSource ="FamilyID"

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =4260
                    Height =300
                    ColumnWidth =2280
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="Text22"
                    ControlSource ="NewProductCode"

                    LayoutCachedLeft =6540
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12900
                    Top =60
                    Width =2220
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="Text25"
                    ControlSource ="NewProductRelease"

                    LayoutCachedLeft =12900
                    LayoutCachedTop =60
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =17520
                    Width =426
                    Height =426
                    Name ="btnFind"
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

                    LayoutCachedLeft =17520
                    LayoutCachedWidth =17946
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
                    Left =10920
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="Text53"
                    ControlSource ="EngineeringChange"

                    LayoutCachedLeft =10920
                    LayoutCachedTop =60
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =5400
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txFamilyDescription"
                    ControlSource ="FamilyDescription"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =1080
                    LayoutCachedTop =60
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15180
                    Top =60
                    Width =2220
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text86"
                    ControlSource ="DateCreated"
                    Format ="Short Date"

                    LayoutCachedLeft =15180
                    LayoutCachedTop =60
                    LayoutCachedWidth =17400
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =18180
                    Width =426
                    Height =426
                    TabIndex =7
                    Name ="btnDeleteRecord"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddddddd177ddddd77dd1ddd1177dddd17dd11dd7117ddd71ddd ,
                        0x111dd1177d117ddd1111d7117711dddd11111d11111ddddd1111dd71117ddddd ,
                        0x111d77111177dddd11d711dd71177ddd1dddddddd71177ddddddddddddd11ddd ,
                        0xdddddddddddddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =18180
                    LayoutCachedWidth =18606
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =1500
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =660
                    Width =1320
                    Name ="btnFiberClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2006
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2006
                    LayoutCachedTop =660
                    LayoutCachedWidth =3326
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3712
                    Top =660
                    Width =1320
                    TabIndex =2
                    Name ="btnAdd"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3712
                    LayoutCachedTop =660
                    LayoutCachedWidth =5032
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5418
                    Top =660
                    Width =1320
                    TabIndex =3
                    Name ="btnsave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5418
                    LayoutCachedTop =660
                    LayoutCachedWidth =6738
                    LayoutCachedHeight =1020
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
Public Event Selected(value As String)

Public Sub btnAdd_Click()
Me.AllowAdditions = True
DoCmd.GoToRecord , , acNewRec
'While Me.CurrentRecord < Me.Recordset.RecordCount
'
'    DoCmd.GoToRecord Record:=acNext
'Wend
Call btnEdit_Click
End Sub


Private Sub btnAddDocs_Click()

DoCmd.OpenForm "frmCableFamilyDocument"
End Sub

Private Sub btnDeleteRecord_Click()
If MsgBox("Are you SURE you want to permantently delete this record?", vbYesNo, "Delete Record") = vbYes Then

    DoCmd.RunSQL "DELETE FROM tblCableFamily where FamilyID = " & Me.FamilyID
    Me.Requery

End If
End Sub

Private Sub btnFiberClose_Click()
DoCmd.Close acForm, Me.name

End Sub



Private Sub btnFind_Click()
    
    DoCmd.OpenForm "frmFindConstruction", acNormal, , "FamilyID = " & Me.FamilyID, , acWindowNormal
End Sub

Private Sub btnRemoveFilter_Click()
Me.ItemNoSearch.value = Null
Me.ItemNoSearch.Requery
Me.Filter = ""
Me.OrderBy = ""
Me.FilterOn = False
Me.OrderByOn = False

'Call Refresh_Click


End Sub

Private Sub btnSave_Click()

Me.AllowAdditions = False
Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If ctl.Locked <> True Then
            ctl.Locked = True
        End If
    End If
Next ctl

Me.ItemNoSearch.Locked = False

End Sub

Private Sub btnSearch_Click()
Dim strFilter As String
strFilter = "[FamilyID] like ""*" & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [FamilyDescription] like   ""*" & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [NewProductCode] like  ""*" & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [EngineeringChange] like   ""*" & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [NewProductRelease] like   ""*" & Me.ItemNoSearch & "*"""

Me.Filter = strFilter
Me.FilterOn = True
End Sub
'Call Refresh_Click



Private Sub btnSortDescending_Click()

Me.OrderBy = "[DateCreated] DESC"
Me.OrderByOn = True

'Call Refresh_Click

End Sub

Private Sub Form_Open(Cancel As Integer)
Me.btnSearch.SetFocus
Me.AllowAdditions = False


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
Private Sub btnEdit_Click()
Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If ctl.Locked <> False Then
            ctl.Locked = False
        End If
    End If
Next ctl
End Sub

Private Sub txFamilyDescription_DblClick(Cancel As Integer)


DoCmd.OpenForm "frmCableFamilyDocument"

RaiseEvent Selected(Me.FamilyID)

End Sub

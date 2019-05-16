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
    Width =11040
    DatasheetFontHeight =10
    ItemSuffix =85
    Left =8580
    Top =1935
    Right =19620
    Bottom =8280
    DatasheetGridlinesColor =12632256
    Filter ="[ListingDocument] like \"*\""
    OrderBy ="[tblBurnListingConstruction].[ListingDocument], [tblBurnListingConstruction].[Co"
        "nstructionID] DESC, [tblBurnListingConstruction].[ListingCompany]"
    RecSrcDt = Begin
        0x8ccd52237622e540
    End
    RecordSource ="tblBurnListingConstruction"
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
        Begin ListBox
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
        Begin FormHeader
            Height =2640
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =2340
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Company"
                    LayoutCachedLeft =240
                    LayoutCachedTop =2340
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =2640
                End
                Begin Label
                    OverlapFlags =85
                    Left =2160
                    Top =2340
                    Width =2760
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Listing Document"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =2340
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =2640
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =180
                    Top =420
                    Width =7440
                    Name ="Box35"
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =1140
                End
                Begin Label
                    OverlapFlags =85
                    Left =6240
                    Top =2340
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="Construction"
                    LayoutCachedLeft =6240
                    LayoutCachedTop =2340
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =2640
                End
                Begin Label
                    OverlapFlags =85
                    Left =4980
                    Top =2340
                    Width =1200
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Section"
                    LayoutCachedLeft =4980
                    LayoutCachedTop =2340
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =2640
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1920
                    Width =10860
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1920
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =1920
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5820
                    Top =1380
                    Width =1320
                    TabIndex =2
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =1740
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
                    Width =7155
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Burn Listing - Add Section/Construction"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =300
                    LayoutCachedTop =540
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =1035
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =4320
                    Top =1380
                    Width =1320
                    TabIndex =1
                    Name ="btnFilter"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =1380
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =1740
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
                    Left =480
                    Top =1380
                    Width =3600
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =480
                    LayoutCachedTop =1380
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1680
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7320
                    Top =1380
                    Width =480
                    TabIndex =3
                    Name ="btnrefresh"
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

                    LayoutCachedLeft =7320
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =1740
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =8340
                    Width =486
                    Height =426
                    Name ="btnFind"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000001e0000001c0000000100180000000000100a0000c40e0000c40e0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffefefefefefefefefefefefefefefefefefefefefefefe ,
                        0xfefefefefefefefefefefefefefefefefefefeffffffffffffffffffffffffff ,
                        0xffff0000fffffffffffffffffffffffffffffffffffffffffffdfdfd9a9a9aff ,
                        0xfffffdfdfdfefefefefefefefefefefefefefefefefefefefefefefefefefefe ,
                        0xfefefefefefefefefefefefefdfdfdfffffff3f3f3adadadf4f4f4ffffff0000 ,
                        0xfffffffffffffffffffffffffffffffffffffffffff9f9f9121212fcfcfcfefe ,
                        0xfef7f7f7f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8f8f9f9f9ffffffe1e1e13d3d3de4e4e4ffffff0000ffffffff ,
                        0xfffffffffffffffffffffffffffffffffffbfbfb2b2b2bfdfdfdffffffd4d4d4 ,
                        0xd9d9d9dadadadadadadadadadadadadadadadadadadadadadadadadadadadada ,
                        0xdad7d7d7e2e2e2fefefee7e7e7515151e8e8e8ffffff0000ffffffffffffffff ,
                        0xfffffffffffffffffffffffffffafafb29292afdfdfdffffffcbccccd3d3d4d3 ,
                        0xd3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d0d0d1 ,
                        0xdbdbdcfefefee7e7e7505050e8e8e8ffffff0000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff363421ffffffffffffe3e0cee9e6d6e8e7d6e8e7 ,
                        0xd6e8e7d6e8e7d6e8e7d6e8e7d6e8e7d6e8e7d6e8e7d6e9e7d6e7e5d3eceadeff ,
                        0xfffee7e7e7505050e8e8e8ffffff0000ffffffecedf7939bd9828cd58891d788 ,
                        0x91d68791de6974ce0002588c96dc8a92d87178b9747cbe757cbe747cbe747cbe ,
                        0x747cbe747cbe747cbe747cbe747cbe747cbe757cbe6b73ba9298ccf5f5fce9e9 ,
                        0xe8505050e8e8e8ffffff0000ffffffd6d6f01e29b10007a50512a9040fad0008 ,
                        0x87acb0d69ba3dc0607a70411a90310a90311aa0412aa0413aa0412aa0411aa04 ,
                        0x11aa0411aa0412aa0412aa0310aa0311aa0003a44549bce7eaf9ebebe8505050 ,
                        0xe8e8e8ffffff0000ffffffd9d8f11c2ab10008a60006a70000ab000579dadbe1 ,
                        0xd4d9f100009c0000a20012af0012aa0004a60000a30008ad0001a80009a7000a ,
                        0xad0000a30008a60216ab0115aa0004a54347bce8ebfaebebe8505050e8e8e8ff ,
                        0xffff0000ffffffd8daf21d2cb90008ad404fc0535fd2000a7cd0d1ddecedfc60 ,
                        0x6ed46d7bd8000493010db1424fc27884d92b36a04e5aba2c3dc7212fa47481d6 ,
                        0x2e3fc0000cb00519b50008ae464ac3e9ebfbebebe8505050e8e8e8ffffff0000 ,
                        0xffffffd9daf11d2bc00008ad8f98d2caccf114197dc8c8dcffffffbbbac2ffff ,
                        0xff9da0c73743bcffffffbfbee19c9abefffff76c78e3676b93ffffff939be300 ,
                        0x00b0051ab90007b4454bc7e9ecfcebebe8505050e8e8e8ffffff0000ffffffd9 ,
                        0xd8f31e2fc2000cba1023bf0e22c9000684d5d6e0d0d7f8000094dee2f2fafafb ,
                        0x7478bbfbfaf53a3db16466befffff75d6ce34e5594ffffff7b85dd0000b3061b ,
                        0xbe0009b9454bcae9ecfbebebe8505050e8e8e8ffffff0000ffffffd9d9f41e2f ,
                        0xc30010bd0115bd000ec5000a85d9d9e0ecf0fe4746c3f3f6fad5d8ef616bc4ff ,
                        0xfff97d83c48e93c5fffffa7781cd9092b6ffffffb0b5e50514bb0115bf000aba ,
                        0x454ccce9ecfbebebe8505050e8e8e8ffffff0000ffffffd9daf31e2fc70010c1 ,
                        0x061cc30117ca010e97afb3ddd5d9f3d4dae9e0e4fb3c41ca1e2fc9bec4ead1d5 ,
                        0xe9dcdcd9ffffffa3a6c1c7c7d5ffffffffffff3f4dca000bbf000bbf454ccfe9 ,
                        0xecfcebebe8505050e8e8e8ffffff0000ffffffd9d9f31e30cb0011c5051cc703 ,
                        0x1ac80117bf2d3bce2a39cc2030c63946d30000bd0014c41628c6303dd35b65c6 ,
                        0xa2a6da515fd8242ea9a4a9e3adb1dd212fc20012c5000cc4464dd3e9ecfcebeb ,
                        0xe7505050e8e8e8ffffff0000ffffffdad8f31e31cd0013c7051dc9051cc9041c ,
                        0xcb0013c70012c70013c70010c5051eca041cca0015c80010c60311bf0f12b804 ,
                        0x12c3000dc40f18b11f2a9d0010be021ccb000ec6464fd5eaecfdebebe7505050 ,
                        0xe8e8e8ffffff0000ffffffd4d7f21022ca0000c4000ac6000ac6000ac6000cc7 ,
                        0x0017d4000ac7000ac6000dca000ecb000fcb000fcb000cca0006c8000ac9000f ,
                        0xcd0008c60008c3000cca000fca0000c73d43d5e6eafceaeae7515151e8e8e8ff ,
                        0xffff0000ffffffe5e4f86b78df5562da5c6adc5b69dc5d6bde5a68de0e1d9265 ,
                        0x72e75f6de04150c44652c94652c94652c94653c94653ca4653c94653c95361d8 ,
                        0x5462da4f5dd3505dd44552cf7c86e1fbfdfffffefc414141e3e3e3ffffff0000 ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffe36331efffffdffff ,
                        0xffd5d4bfdddac8dddbc7dddbc7dddbc7dddbc7dedbc8e6e4d19b98849e9b87b4 ,
                        0xb39fb0ae9bb0ac99afad9ebcbcb96e6e6ea5a5a5fdfdfdffffff0000ffffffff ,
                        0xfffffffffffffffffffffffffffffffffffbfbfb292a2afefefdffffffc0c0c0 ,
                        0xc8c8c9c8c8c9c8c8c9c8c8c9c8c8c9c9c9cad8d8d95253535454557777787070 ,
                        0x717171727575757b7b7b111111c6c6c6ffffffffffff0000ffffffffffffffff ,
                        0xfffffffffffffffffffffffffffbfbfb2a2a2afdfdfdffffffc7c7c7cccccccd ,
                        0xcdcdcdcdcdcdcdcdcdcdcdcececedadada656565ecececffffffffffffffffff ,
                        0xffffff737373bebebeffffffffffffffffff0000ffffffffffffffffffffffff ,
                        0xfffffffffffffffffffbfbfb2a2a2afdfdfdffffffcbcbcbd2d2d2d2d2d2d2d2 ,
                        0xd2d2d2d2d2d2d2d3d3d3dfdfdf646464d5d5d5fffffff5f5f5bfbfbfa3a3a3e1 ,
                        0xe1e1ffffffffffffffffffffffff0000ffffffffffffffffffffffffffffffff ,
                        0xfffffffffffbfbfb2a2a2afdfdfdffffffd0d0d0d5d5d5d5d5d5d5d5d5d5d5d5 ,
                        0xd5d5d5d6d6d6e2e2e2656565d6d6d6ffffffcecece414141f1f1f1ffffffffff ,
                        0xffffffffffffffffffff0000ffffffffffffffffffffffffffffffffffffffff ,
                        0xfffbfbfb2a2a2afdfdfdffffffd3d3d3d8d8d8d8d8d8d8d8d8d8d8d8d8d8d8d9 ,
                        0xd9d9e6e6e6696969e2e2e2fefefe787878a8a8a8ffffffffffffffffffffffff ,
                        0xffffffffffff0000fffffffffffffffffffffffffffffffffffffffffffbfbfb ,
                        0x212121fffffffffffff5f5f5f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7ffff ,
                        0xff7f7f7f5c5c5cc0c0c0c2c2c2ffffffffffffffffffffffffffffffffffffff ,
                        0xffff0000fffffffffffffffffffffffffffffffffffffffffffdfdfd4d4d4df1 ,
                        0xf1f1ebebebececececececececececececececececececebebebf9f9f97a7a7a ,
                        0x818181f2f2f2ffffffffffffffffffffffffffffffffffffffffffffffff0000 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffa0a0a0a2a2a2a1a1 ,
                        0xa1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a2a2a29f9f9fa9a9a9ffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffff0000ffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffff9f9f9f7f7f7f7f7f7f7f7f7 ,
                        0xf7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f8f8f8f7f7f7fafafaffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff0000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="pdf_file.bmp"

                    LayoutCachedLeft =8340
                    LayoutCachedWidth =8826
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbListingCompany"
                    ControlSource ="ListingCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="tblBurnListingCompany"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =360
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =60
                    Width =2760
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbListingDocument"
                    ControlSource ="ListingDocument"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ListingDocument FROM tblBurnListingDocument WHERE ListingCompany = 'ETL'"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2160
                    LayoutCachedTop =60
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4980
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txSection"
                    ControlSource ="ListingSection"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =60
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6240
                    Top =60
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="txConstruction"
                    ControlSource ="ListingConstruction"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =60
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9000
                    Width =486
                    Height =426
                    TabIndex =5
                    Name ="btnAddSafetyRating"
                    Caption ="btnAddSafetyRating"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000c40e0000c40e0000 ,
                        0x0000000000000000fffffff7f7f78585852b2b2b020202000000000000000000 ,
                        0x000000000000000000000000000000000000000000404040ffffffffffffffff ,
                        0xffffffff39393900000000000000000000000000000000000000000000000000 ,
                        0x00000000000202022c2c2c888888f9f9f9fffffff4f4f43d3d3d020202010101 ,
                        0x0101010000000000000000000000000000000000000000000000000000000000 ,
                        0x00404040ffffffffffffffffffffffff39393900000000000000000000000000 ,
                        0x0000000000000000000000000000000000010101010101020202434343f9f9f9 ,
                        0x8585850101010101010000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000404040ffffffffffffffffffffffff39393900 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000101010101018888882c2c2c0000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000404040ffffffff ,
                        0xffffffffffffffff393939000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000002c2c2c0202020000000000 ,
                        0x000000000000002929297f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f ,
                        0x7f7f7f7f7f9f9f9fffffffffffffffffffffffff9c9c9c7f7f7f7f7f7f7f7f7f ,
                        0x7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f2525250000000000000000000101 ,
                        0x010202020000000000000000000000002c2c2cf9f9f9ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff5f5 ,
                        0xf52424240000000000000000000000000000000000000000000000007f7f7fff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f00000000000000000000000000000000 ,
                        0x00000000000000007f7f7fffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f00000000 ,
                        0x00000000000000000000000000000000000000007f7f7fffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffff7f7f7f000000000000000000000000000000000000000000000000 ,
                        0x7f7f7fffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffbfbfbf000000000000000000000000c6c6c6ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffff7f7f7f000000000000000000000000 ,
                        0x0000000000000000000000007f7f7fffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffbfbfbf000000000000000000000000c6c6c6ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x0000000000000000000000000000000000000000000000007f7f7fffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffbfbfbf00000000 ,
                        0x0000000000000000c6c6c6ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff7f7f7f0000000000000000000000000000000000000000 ,
                        0x000000007f7f7fffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffbfbfbf000000000000000000000000c6c6c6ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff7f7f7f0000000000000000 ,
                        0x000000000000000000000000000000007f7f7fffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffbfbfbf000000000000000000000000 ,
                        0xc6c6c6ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xff7f7f7f0000000000000000000000000000000000000000000000007f7f7fff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffbfbfbf ,
                        0x000000000000000000000000c6c6c6ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f00000000000000000000000039393939 ,
                        0x39393939393939399c9c9cffffffffffffffffffffffffc6c6c6c6c6c6c6c6c6 ,
                        0xc6c6c6c6c6c6c6c6c69595950000000000000000000000009a9a9ac6c6c6c6c6 ,
                        0xc6c6c6c6c6c6c6c6c6c6c6c6c6ffffffffffffffffffffffff9c9c9c39393939 ,
                        0x3939393939393939ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffff0000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffff0000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000ffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000ffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff4040404040404040 ,
                        0x404040409f9f9fffffffffffffffffffffffffbfbfbfbfbfbfbfbfbfbfbfbfbf ,
                        0xbfbfbfbfbf8f8f8f000000000000000000000000959595bfbfbfbfbfbfbfbfbf ,
                        0xbfbfbfbfbfbfbfbfbfffffffffffffffffffffffff9f9f9f4040404040404040 ,
                        0x404040400000000000000000000000007f7f7fffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffbfbfbf000000000000000000000000 ,
                        0xc6c6c6ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xff7f7f7f0000000000000000000000000000000000000000000000007f7f7fff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffbfbfbf ,
                        0x000000000000000000000000c6c6c6ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f00000000000000000000000000000000 ,
                        0x00000000000000007f7f7fffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffbfbfbf000000000000000000000000c6c6c6ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f00000000 ,
                        0x00000000000000000000000000000000000000007f7f7fffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffbfbfbf0000000000000000 ,
                        0x00000000c6c6c6ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffff7f7f7f000000000000000000000000000000000000000000000000 ,
                        0x7f7f7fffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffbfbfbf000000000000000000000000c6c6c6ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffff7f7f7f000000000000000000000000 ,
                        0x0000000000000000000000007f7f7fffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffbfbfbf000000000000000000000000c6c6c6ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x0000000000000000000000000000000000000000000000007f7f7fffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff7f7f7f0000000000000000000000000000000000000000 ,
                        0x000000007f7f7fffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffff7f7f7f0000000000000000 ,
                        0x000000000000000000000000000000007f7f7fffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xff7f7f7f0000000000000000000000000000000000000000000000002f2f2ffd ,
                        0xfdfdffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffff9f9f928282800000000000000000000000002020201 ,
                        0x01010000000000000000002f2f2f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f ,
                        0x7f7f7f7f7f7f7f7f7f9f9f9fffffffffffffffffffffffff9c9c9c7f7f7f7f7f ,
                        0x7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f2c2c2c00000000000000 ,
                        0x00000101010202022b2b2b000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000404040ffffffffffffffff ,
                        0xffffffff39393900000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000002b2b2b828282010101000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00404040ffffffffffffffffffffffff39393900000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000010101848484 ,
                        0xf4f4f43737370202020101010101010000000000000000000000000000000000 ,
                        0x00000000000000000000000000404040ffffffffffffffffffffffff39393900 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0101010202023d3d3df7f7f7fffffff4f4f48282822b2b2b0202020000000000 ,
                        0x00000000000000000000000000000000000000000000000000404040ffffffff ,
                        0xffffffffffffffff393939000000000000000000000000000000000000000000 ,
                        0x0000000000000000000202022c2c2c868686f4f4f4ffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add Safety Rating"
                    Picture ="add_318-70453.bmp"

                    LayoutCachedLeft =9000
                    LayoutCachedWidth =9486
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =9780
                    Top =60
                    Width =480
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="txConstructionID"
                    ControlSource ="ConstructionID"

                    LayoutCachedLeft =9780
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =9780
                    Top =60
                    Width =480
                    Height =300
                    TabIndex =7
                    Name ="btnDelete"
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

                    LayoutCachedLeft =9780
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10440
                    Top =60
                    Width =480
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="Text84"
                    ControlSource ="ConstructionID"

                    LayoutCachedLeft =10440
                    LayoutCachedTop =60
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =1200
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =660
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =300
                    LayoutCachedTop =660
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2020
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2020
                    LayoutCachedTop =660
                    LayoutCachedWidth =3340
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3740
                    Top =660
                    Width =1320
                    TabIndex =2
                    Name ="btnAdd"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3740
                    LayoutCachedTop =660
                    LayoutCachedWidth =5060
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5460
                    Top =660
                    Width =1320
                    TabIndex =3
                    Name ="btnsave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5460
                    LayoutCachedTop =660
                    LayoutCachedWidth =6780
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
Option Explicit
Option Compare Database

Private Sub btnAdd_Click()
Me.AllowAdditions = True
DoCmd.GoToRecord , , acNewRec
Call btnEdit_Click
End Sub


Private Sub btnClose_Click()
DoCmd.Close acForm, Me.name
End Sub

Private Sub btnDelete_Click()

   If MsgBox("Are you SURE you want to delete this record?" & vbCrLf & _
    "This will permanently delete the record.", vbYesNo, "2nd Delete Confirmation") = vbYes Then
         DoCmd.RunSQL "DELETE FROM tblBurnListingConstruction WHERE ConstructionID = " & Me.ConstructionID
         Me.Requery
   End If

End Sub

Private Sub btnEdit_Click()
Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
        If ctl.Locked <> False Then
            ctl.Locked = False
        End If
    End If
Next ctl
End Sub

Private Sub btnFilter_Click()
Dim strFilter As String

strFilter = "[ListingDocument] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [ListingCompany] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [ListingSection] like """ & Me.ItemNoSearch & "*"""

Me.Filter = strFilter
Me.FilterOn = True
Call btnRefresh_Click

End Sub

Private Sub btnFind_Click()
On Error GoTo ErrorHandler:
'DoCmd.OpenForm "frmFindBurnListing_Print", acNormal, , "Tab = '" & Me.ListingTab & "' AND Section = '" & Me.ListingSection & "' AND Construction = '" & Me.ListingConstruction & "'", acFormReadOnly
Call FindBurnListing(Me.ListingCompany, Me.ListingDocument)

ErrorExit:
Exit Sub

ErrorHandler:
MsgBox ("Missing listing information for base")
Resume ErrorExit:


End Sub

Private Sub btnRefresh_Click()
Me.Requery
Me.Refresh
End Sub

Private Sub btnRemoveFilter_Click()
Me.ItemNoSearch.value = Null
Me.ItemNoSearch.Requery
Me.Filter = ""
Me.FilterOn = False
'Call Refresh_Click
End Sub
Private Sub btnSave_Click()
'If Me.Type Like "ft*" Then
'    Me.Canada = -1
'    Me.US = 0
'Else
'    Me.Canada = 0
'    Me.US = -1
'End If
Me.AllowAdditions = False
Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
        If ctl.Locked <> True Then
            ctl.Locked = True
        End If
    End If
Next ctl

End Sub



Private Sub cbListingCompany_AfterUpdate()
Call FilterListingDocument
Me.cbListingDocument = ""
Me.cbListingDocument.Requery

End Sub

Private Sub cbListingDocument_AfterUpdate()
Call FilterListingDocument
End Sub

Private Sub btnAddSafetyRating_Click()
If Me.Dirty Then Me.Dirty = False
DoCmd.OpenForm "frmBurnListingAddRating", acNormal, , "ConstructionID = " & Me.txConstructionID
Forms!frmBurnListingAddRating!BurnListingConstructionID = Me.txConstructionID
End Sub

Private Sub Form_Open(Cancel As Integer)
Me.Filter = "[ListingDocument] like """ & Me.ItemNoSearch & "*"""
Me.FilterOn = True
Me.AllowAdditions = False
Call btnRefresh_Click

End Sub

Private Sub ItemNoSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.btnFilter.SetFocus
        Call btnFilter_Click
    End If
End Sub

Private Sub FilterListingDocument()

Dim ListingDocSource As String

    ListingDocSource = "SELECT ListingDocument " & _
                     "FROM tblBurnListingDocument " & _
                     "WHERE ListingCompany = '" & Me.cbListingCompany.value & "'"
    Me.cbListingDocument.rowSource = ListingDocSource
    
'    ListingSecSource = "SELECT ListingSection " & _
'                     "FROM tblBurnListingConstruction " & _
'                     "WHERE ListingCompany = '" & Me.cbListingCompany.Value & "' AND ListingDocument = '" & Me.cbListingDocument & "'"
'    Me.cbSection.RowSource = ListingSecSource
'
    
    'Me.cbListingCompany.Requery


End Sub

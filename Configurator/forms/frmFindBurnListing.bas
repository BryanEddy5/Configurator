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
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13860
    DatasheetFontHeight =10
    ItemSuffix =74
    Left =5235
    Top =885
    Right =19095
    Bottom =12390
    DatasheetGridlinesColor =12632256
    Filter ="[Base] like \"l**\" OR [ListingCompany] like \"l**\" OR [ListingTAb] like \"l**\""
    OrderBy ="[frmFindBurnListing].[Base]"
    RecSrcDt = Begin
        0x03986d09eefee440
    End
    RecordSource ="SELECT DISTINCT tblBurnListing.Base, tblBurnListing.ListingTab, tblBurnListing.L"
        "istingSection, tblBurnListing.ListingConstruction, tblBurnListing.Field6, tblBur"
        "nListingRating.Rating, tblBurnListingConstruction.ListingCompany, tblCableConstr"
        "uctions.CablePasses FROM (tblCableConstructionReferences INNER JOIN ((tblBurnLis"
        "tingConstruction INNER JOIN tblBurnListing ON (tblBurnListingConstruction.Listin"
        "gCompany = tblBurnListing.ListingCompany) AND (tblBurnListingConstruction.Listin"
        "gDocument = tblBurnListing.ListingTab) AND (tblBurnListingConstruction.ListingSe"
        "ction = tblBurnListing.ListingSection) AND (tblBurnListingConstruction.ListingCo"
        "nstruction = tblBurnListing.ListingConstruction)) INNER JOIN tblBurnListingRatin"
        "g ON tblBurnListingConstruction.ConstructionID = tblBurnListingRating.Constructi"
        "onID) ON tblCableConstructionReferences.Base = tblBurnListing.Base) LEFT JOIN tb"
        "lCableConstructions ON tblCableConstructionReferences.BaseID = tblCableConstruct"
        "ions.BaseID; "
    Caption ="Item Search"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
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
            Height =2040
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Base"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =9420
                    Top =1680
                    Width =1140
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Construction"
                    LayoutCachedLeft =9420
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =4020
                    Top =1680
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Listing Company"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5580
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
                    Left =10740
                    Top =1680
                    Width =990
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="Comments"
                    LayoutCachedLeft =10740
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11730
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =8340
                    Top =1680
                    Width =900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="Section"
                    LayoutCachedLeft =8340
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =5640
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Tab"
                    LayoutCachedLeft =5640
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1260
                    Width =13440
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =1260
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8520
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"

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
                    Caption ="Burn Listing"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =7020
                    Top =660
                    Width =1320
                    TabIndex =2
                    Name ="btnFilter"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"

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
                    Top =660
                    Width =3600
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    Left =1980
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label68"
                    Caption ="Cable Rating"
                    LayoutCachedLeft =1980
                    LayoutCachedTop =1680
                    LayoutCachedWidth =3840
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="Base"
                    ControlSource ="Base"

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9420
                    Top =60
                    Width =1140
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="ListingConstruction"
                    ControlSource ="ListingConstruction"

                    LayoutCachedLeft =9420
                    LayoutCachedTop =60
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4020
                    Top =60
                    Width =1560
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="ListingCompany"
                    ControlSource ="ListingCompany"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =60
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10740
                    Top =60
                    Width =2280
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="Comments"
                    ControlSource ="Field6"

                    LayoutCachedLeft =10740
                    LayoutCachedTop =60
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8340
                    Top =60
                    Width =960
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="ListingSection"
                    ControlSource ="ListingSection"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5640
                    Top =60
                    Width =2100
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="ListingTab"
                    ControlSource ="ListingTab"

                    LayoutCachedLeft =5640
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1980
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text69"
                    ControlSource ="Rating"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7800
                    Width =486
                    Height =426
                    TabIndex =7
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

                    LayoutCachedLeft =7800
                    LayoutCachedWidth =8286
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =13080
                    Top =60
                    Width =600
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="txLocation"
                    ControlSource ="CablePasses"

                    LayoutCachedLeft =13080
                    LayoutCachedTop =60
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =13080
                    Width =426
                    Height =426
                    TabIndex =8
                    Name ="btnConstruction"
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

                    LayoutCachedLeft =13080
                    LayoutCachedWidth =13506
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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
                    Name ="btnClose"
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
                    Left =1920
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="Refresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3540
                    Top =660
                    Width =1620
                    TabIndex =2
                    Name ="Command70"
                    Caption ="Update Listing File"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =660
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5400
                    Top =660
                    Width =1620
                    TabIndex =3
                    Name ="btnBurnListingDocument"
                    Caption ="View All Listings"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5400
                    LayoutCachedTop =660
                    LayoutCachedWidth =7020
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
    'While Me.CurrentRecord < Me.Recordset.RecordCount
    '
    '    DoCmd.GoToRecord Record:=acNext
    'Wend
    Call btnEdit_Click
End Sub

Private Sub btnClose_Click()
    DoCmd.Close
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

Private Sub btnConstruction_Click()
    Call OpenAClient(3, Forms!frmFindBurnListing!Base, Nz(Me.CablePasses, 2))
End Sub

Private Sub btnFilter_Click()
    Dim strFilter As String

    strFilter = "[Base] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [ListingCompany] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [ListingTAb] like """ & Me.ItemNoSearch & "*"""

    Me.Filter = strFilter
    Me.FilterOn = True
    Call Refresh_Click

End Sub

Private Sub btnFind_Click()
    On Error GoTo ErrorHandler:
    'DoCmd.OpenForm "frmFindBurnListing_Print", acNormal, , "Tab = '" & Me.ListingTab & "' AND Section = '" & Me.ListingSection & "' AND Construction = '" & Me.ListingConstruction & "'", acFormReadOnly
    FindBurnListing Me.ListingCompany, Me.ListingTab

ErrorExit:
    Exit Sub

ErrorHandler:
    MsgBox ("Missing listing information for base")
    Resume ErrorExit:


End Sub

Private Sub btnRemoveFilter_Click()
    Me.ItemNoSearch.value = Null
    Me.ItemNoSearch.Requery
    Me.Filter = ""
    Me.FilterOn = False
    Call Refresh_Click
End Sub

Private Sub Command70_Click()
    DoCmd.OpenForm "frmNewListingUpdate"
    'Forms!frmCustomerSpecsNewRecord.
End Sub

Private Sub btnBurnListingDocument_Click()
    DoCmd.OpenForm "frmFindBurnListingDocument"
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.Filter = "[Base] like """ & Me.ItemNoSearch & "*"""
    Me.FilterOn = True
    Me.AllowAdditions = False
    Call Refresh_Click

End Sub

Private Sub ItemNoSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.btnFilter.SetFocus
        Call btnFilter_Click
    End If
End Sub

Private Sub Refresh_Click()
    Me.Requery
    Me.Refresh
End Sub

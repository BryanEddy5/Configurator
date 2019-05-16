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
    Width =28320
    DatasheetFontHeight =10
    ItemSuffix =104
    Left =1230
    Top =6195
    Right =29460
    Bottom =20295
    DatasheetGridlinesColor =12632256
    Filter ="([Base] like \"br*cu*\" OR [NominalOD] like \"br*cu*\" OR [ConstructionDescripti"
        "on] like \"br*cu*\" OR [DesignCode] like \"br*cu*\" OR [JacketMaterial] like \"b"
        "r*cu*\" OR [StandardOperation] like \"br*cu*\" OR [BaseID] like \"br*cu*\" OR [T"
        "BType] like \"br*cu*\")"
    OrderBy ="[vConstructionSearch].[Active], [vConstructionSearch].[BaseID] DESC, [vConstruct"
        "ionSearch].[DateCreated] DESC"
    RecSrcDt = Begin
        0xd9f2680ed4cbe440
    End
    RecordSource ="vConstructionSearch"
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
                    Left =60
                    Top =1680
                    Width =2100
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Base"
                    LayoutCachedLeft =60
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =20820
                    Top =1680
                    Width =1140
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Height"
                    LayoutCachedLeft =20820
                    LayoutCachedTop =1680
                    LayoutCachedWidth =21960
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =3060
                    Top =1680
                    Width =1320
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Design Code"
                    LayoutCachedLeft =3060
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =14160
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label24"
                    Caption ="Active"
                    LayoutCachedLeft =14160
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =180
                    Top =420
                    Width =3660
                    Name ="Box35"
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =1140
                End
                Begin Label
                    OverlapFlags =85
                    Left =22020
                    Top =1680
                    Width =1320
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label43"
                    Caption ="Jacket"
                    LayoutCachedLeft =22020
                    LayoutCachedTop =1680
                    LayoutCachedWidth =23340
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =17520
                    Top =1680
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="Nom OD"
                    LayoutCachedLeft =17520
                    LayoutCachedTop =1680
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =27540
                    Top =1680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label45"
                    Caption ="Passes"
                    LayoutCachedLeft =27540
                    LayoutCachedTop =1680
                    LayoutCachedWidth =28320
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1260
                    Width =26580
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =26760
                    LayoutCachedHeight =1260
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8500
                    Top =780
                    Width =450
                    Height =420
                    TabIndex =3
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

                    LayoutCachedLeft =8500
                    LayoutCachedTop =780
                    LayoutCachedWidth =8950
                    LayoutCachedHeight =1200
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
                    Width =3420
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Construction Search"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =300
                    LayoutCachedTop =540
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    Left =23400
                    Top =1680
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label61"
                    Caption ="Setup"
                    LayoutCachedLeft =23400
                    LayoutCachedTop =1680
                    LayoutCachedWidth =24840
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =7860
                    Top =780
                    Width =450
                    Height =420
                    TabIndex =2
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

                    LayoutCachedLeft =7860
                    LayoutCachedTop =780
                    LayoutCachedWidth =8310
                    LayoutCachedHeight =1200
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
                    Left =4140
                    Top =780
                    Width =3600
                    Height =300
                    ColumnOrder =2
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =780
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    Left =4080
                    Top =60
                    Width =3600
                    Height =600
                    FontSize =9
                    Name ="Label65"
                    Caption ="Master Search: Searches all fields displayed with embedded wild card."
                    LayoutCachedLeft =4080
                    LayoutCachedTop =60
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =660
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9140
                    Top =780
                    Width =450
                    Height =420
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
                    ControlTipText ="Refresh"

                    LayoutCachedLeft =9140
                    LayoutCachedTop =780
                    LayoutCachedWidth =9590
                    LayoutCachedHeight =1200
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =24900
                    Top =1680
                    Width =1710
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label66"
                    Caption ="Fiber Configuration"
                    LayoutCachedLeft =24900
                    LayoutCachedTop =1680
                    LayoutCachedWidth =26610
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =2220
                    Top =1680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label68"
                    Caption ="BaseID"
                    LayoutCachedLeft =2220
                    LayoutCachedTop =1680
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =15780
                    Top =1680
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label70"
                    Caption ="Reason"
                    LayoutCachedLeft =15780
                    LayoutCachedTop =1680
                    LayoutCachedWidth =17460
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =18420
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label73"
                    Caption ="SubOD"
                    LayoutCachedLeft =18420
                    LayoutCachedTop =1680
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =19140
                    Top =1680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label77"
                    Caption ="# Subs"
                    LayoutCachedLeft =19140
                    LayoutCachedTop =1680
                    LayoutCachedWidth =19920
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =19980
                    Top =1680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label79"
                    Caption ="# Fillers"
                    LayoutCachedLeft =19980
                    LayoutCachedTop =1680
                    LayoutCachedWidth =20760
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =14880
                    Top =1680
                    Width =855
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label81"
                    Caption ="Released"
                    LayoutCachedLeft =14880
                    LayoutCachedTop =1680
                    LayoutCachedWidth =15735
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =26640
                    Top =1680
                    Width =810
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label83"
                    Caption ="FamilyID"
                    LayoutCachedLeft =26640
                    LayoutCachedTop =1680
                    LayoutCachedWidth =27450
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9780
                    Top =780
                    Width =450
                    Height =420
                    TabIndex =5
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
                    ControlTipText ="Export to Excel"

                    LayoutCachedLeft =9780
                    LayoutCachedTop =780
                    LayoutCachedWidth =10230
                    LayoutCachedHeight =1200
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =12000
                    Top =1680
                    Width =1200
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label86"
                    Caption ="#F per Sub"
                    LayoutCachedLeft =12000
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =6300
                    Top =1680
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label88"
                    Caption ="Family Desc"
                    LayoutCachedLeft =6300
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1980
                End
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =12600
                    Top =840
                    Width =4695
                    Height =285
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =6
                    ColumnInfo ="\"\";\"\";\"10\";\"140\""
                    Name ="cbFamilyDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblDesignTypes.SalesForceDescription, tblDesignTypes.Active FROM"
                        " tblDesignTypes ORDER BY tblDesignTypes.SalesForceDescription; "
                    AfterUpdate ="[Event Procedure]"
                    Format ="@;\" <search by family>\""

                    LayoutCachedLeft =12600
                    LayoutCachedTop =840
                    LayoutCachedWidth =17295
                    LayoutCachedHeight =1125
                End
                Begin Label
                    OverlapFlags =85
                    Left =11280
                    Top =1680
                    Width =645
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label92"
                    Caption ="#Fiber"
                    LayoutCachedLeft =11280
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11925
                    LayoutCachedHeight =1980
                End
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =17520
                    Top =840
                    Width =3180
                    Height =285
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =7
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbJacketDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblDesignCodeJacket.JacketDescription_SalesForce, tblDesignTypes"
                        ".SalesForceDescription FROM tblDesignCodeJacket INNER JOIN (tblCableConstruction"
                        "References INNER JOIN tblDesignTypes ON tblCableConstructionReferences.DesignTyp"
                        "eID = tblDesignTypes.DesignTypeID) ON (tblCableConstructionReferences.CableFamil"
                        "y = tblDesignCodeJacket.CableLetter) AND (tblDesignCodeJacket.JacketLetter = tbl"
                        "CableConstructionReferences.CableLevel1) WHERE (((tblDesignTypes.SalesForceDescr"
                        "iption)=[Forms]![frmFindConstruction]![cbFamilyDesc])); "
                    AfterUpdate ="[Event Procedure]"
                    Format ="@;\" <search by jacket>\""

                    LayoutCachedLeft =17520
                    LayoutCachedTop =840
                    LayoutCachedWidth =20700
                    LayoutCachedHeight =1125
                End
                Begin Label
                    OverlapFlags =87
                    Left =12600
                    Top =540
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label95"
                    Caption ="Family Desc"
                    LayoutCachedLeft =12600
                    LayoutCachedTop =540
                    LayoutCachedWidth =16380
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =87
                    Left =17520
                    Top =540
                    Width =1065
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label96"
                    Caption ="Jacket Desc"
                    LayoutCachedLeft =17520
                    LayoutCachedTop =540
                    LayoutCachedWidth =18585
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =85
                    Left =10140
                    Top =1680
                    Width =1065
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label97"
                    Caption ="Jacket Desc"
                    LayoutCachedLeft =10140
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11205
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =13260
                    Top =1680
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label99"
                    Caption ="Quotable"
                    LayoutCachedLeft =13260
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =4860
                    Top =1680
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label102"
                    Caption ="Creation Date"
                    LayoutCachedLeft =4860
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =26340
                    Top =300
                    Width =630
                    Height =630
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
                    ControlTipText ="Close"
                    Picture ="CloseIcon.bmp"

                    LayoutCachedLeft =26340
                    LayoutCachedTop =300
                    LayoutCachedWidth =26970
                    LayoutCachedHeight =930
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
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =2100
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txBase"
                    ControlSource ="Base"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20820
                    Top =60
                    Width =1140
                    Height =300
                    ColumnWidth =2835
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =500
                    TabIndex =17
                    Name ="txHeigh"
                    ControlSource ="Height"

                    LayoutCachedLeft =20820
                    LayoutCachedTop =60
                    LayoutCachedWidth =21960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3060
                    Top =60
                    Width =1320
                    Height =300
                    ColumnWidth =2280
                    ColumnOrder =2
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txDesignCode"
                    ControlSource ="DesignCode"

                    LayoutCachedLeft =3060
                    LayoutCachedTop =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14160
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="txActive"
                    ControlSource ="Active"
                    Format ="Yes/No"

                    LayoutCachedLeft =14160
                    LayoutCachedTop =60
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
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

                    LayoutCachedLeft =4440
                    LayoutCachedWidth =4866
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
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =27540
                    Top =60
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =22
                    Name ="txLocation"
                    ControlSource ="Location"

                    LayoutCachedLeft =27540
                    LayoutCachedTop =60
                    LayoutCachedWidth =27900
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22020
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =18
                    Name ="txJacket"
                    ControlSource ="JacketMaterial"

                    LayoutCachedLeft =22020
                    LayoutCachedTop =60
                    LayoutCachedWidth =23340
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17520
                    Top =60
                    Width =840
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="txNominalOD"
                    ControlSource ="NominalOD"

                    LayoutCachedLeft =17520
                    LayoutCachedTop =60
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =23400
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =19
                    Name ="txSetup"
                    ControlSource ="StandardOperation"

                    LayoutCachedLeft =23400
                    LayoutCachedTop =60
                    LayoutCachedWidth =24840
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2220
                    Top =60
                    Width =780
                    Height =300
                    ColumnOrder =4
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txBaseID"
                    ControlSource ="BaseID"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =60
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =24900
                    Top =60
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =20
                    Name ="txTBType"
                    ControlSource ="TBType"

                    LayoutCachedLeft =24900
                    LayoutCachedTop =60
                    LayoutCachedWidth =26580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15780
                    Top =60
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =12
                    Name ="txInactiveReason"
                    ControlSource ="InactiveReason"

                    LayoutCachedLeft =15780
                    LayoutCachedTop =60
                    LayoutCachedWidth =17460
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18420
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =14
                    Name ="txSubOD"
                    ControlSource ="SubOD"

                    LayoutCachedLeft =18420
                    LayoutCachedTop =60
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19140
                    Top =60
                    Width =780
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =15
                    Name ="txNumSubs"
                    ControlSource ="NumSubPositions"

                    LayoutCachedLeft =19140
                    LayoutCachedTop =60
                    LayoutCachedWidth =19920
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19980
                    Top =60
                    Width =780
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =16
                    Name ="txNumFillers"
                    ControlSource ="NumSubFillers"

                    LayoutCachedLeft =19980
                    LayoutCachedTop =60
                    LayoutCachedWidth =20760
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =26640
                    Top =60
                    Width =840
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =21
                    Name ="txFamilyID"
                    ControlSource ="FamilyID"

                    LayoutCachedLeft =26640
                    LayoutCachedTop =60
                    LayoutCachedWidth =27480
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14880
                    Top =60
                    Width =840
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="txReleasedDesign"
                    ControlSource ="ReleasedDesign"
                    Format ="Yes/No"

                    LayoutCachedLeft =14880
                    LayoutCachedTop =60
                    LayoutCachedWidth =15720
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12000
                    Top =60
                    Width =1200
                    Height =300
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="txSubFiberCount"
                    ControlSource ="SubFiberCount"

                    LayoutCachedLeft =12000
                    LayoutCachedTop =60
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6300
                    Top =60
                    Width =3780
                    Height =300
                    ColumnOrder =6
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="txSalesForceDescription"
                    ControlSource ="SalesForceDescription"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =60
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11280
                    Top =60
                    Width =660
                    Height =300
                    ColumnOrder =8
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="Text93"
                    ControlSource ="FiberCount"

                    LayoutCachedLeft =11280
                    LayoutCachedTop =60
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10140
                    Top =60
                    Width =1080
                    Height =300
                    ColumnOrder =7
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text98"
                    ControlSource ="JacketDescription_SalesForce"

                    LayoutCachedLeft =10140
                    LayoutCachedTop =60
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13260
                    Top =60
                    Width =840
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="Text100"
                    ControlSource ="ApprovedForQuoting"
                    Format ="Yes/No"

                    LayoutCachedLeft =13260
                    LayoutCachedTop =60
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4920
                    Top =60
                    Width =1335
                    Height =300
                    ColumnOrder =5
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="Text101"
                    ControlSource ="DateCreated"
                    Format ="Medium Date"
                    ShowDatePicker =0

                    LayoutCachedLeft =4920
                    LayoutCachedTop =60
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =27840
                    Top =60
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =23
                    Name ="txConstructionID"
                    ControlSource ="ConstructionID"

                    LayoutCachedLeft =27840
                    LayoutCachedTop =60
                    LayoutCachedWidth =28200
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =540
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =300
                    Width =4380
                    Height =240
                    BorderColor =-2147483633
                    ForeColor =4144959
                    Name ="lbMessage"
                    Caption ="| Type to search  | Enter to select  | Esc to cancel"
                    LayoutCachedLeft =120
                    LayoutCachedTop =300
                    LayoutCachedWidth =4500
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

Private Sub btnExcelExport_Click()

    On Error GoTo ErrorHandler

    DoCmd.OutputTo acOutputForm, "frmFindConstruction", acFormatXLS, Environ("USERPROFILE") & "\Desktop\ItemConstructionResults" & Format(Now, "yyyyMMddhhmmss") & ".xls"
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
    Call OpenAClient(2, Me.txConstructionID, Nz(Me.txLocation, 2))
End Sub

Private Sub btnRemoveFilter_Click()
    Me.ItemNoSearch.value = Null
    Me.ItemNoSearch.Requery
    Me.cbFamilyDesc = Null
    Me.cbJacketDesc = Null
    Me.cbJacketDesc.Requery
    Me.Filter = ""
    Me.OrderBy = ""
    Me.FilterOn = False
    Me.OrderByOn = False
    Me.cbFamilyDesc = Null

    'Call Refresh_Click


End Sub

Private Sub btnSearch_Click()
    Dim strFilter As String
    strFilter = "([Base] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [NominalOD] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [ConstructionDescription] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [DesignCode] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [JacketMaterial] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [StandardOperation] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [BaseID] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [TBType] like """ & Me.ItemNoSearch & "*"")"

    If Not IsNull(Me.cbFamilyDesc) Then strFilter = strFilter & " AND [SalesForceDescription] like """ & Me.cbFamilyDesc & "*"""
    If Not IsNull(Me.cbJacketDesc) Then strFilter = strFilter & " AND [JacketDescription_SalesForce]  like """ & Me.cbJacketDesc & "*"""


    Me.Filter = strFilter
    Me.FilterOn = True
    'Call Refresh_Click
End Sub

Private Sub btnSortDescending_Click()

    Me.OrderBy = "[ConstructionDescription] DESC"
    Me.OrderByOn = True

    'Call Refresh_Click

End Sub

Private Sub cbFamilyDesc_AfterUpdate()

    Call btnSearch_Click
    Me.cbJacketDesc = Null
    Me.cbJacketDesc.Requery

End Sub

Private Sub cbJacketDesc_AfterUpdate()
    Call btnSearch_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.ItemNoSearch.SetFocus
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

Private Sub txBase_DblClick(Cancel As Integer)

    Call OpenAClient(FindItem, "base = '" & Me.Base & "'")


End Sub

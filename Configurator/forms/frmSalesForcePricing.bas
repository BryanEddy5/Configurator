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
    AllowDesignChanges = NotDefault
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =24480
    DatasheetFontHeight =10
    ItemSuffix =111
    Left =1905
    Top =3600
    Right =26325
    Bottom =12000
    DatasheetGridlinesColor =12632256
    OrderBy ="[tblSalesForce_Catalog_Pricing].[Multiplier]"
    RecSrcDt = Begin
        0x4ffacb9d4bece440
    End
    RecordSource ="tblSalesForce_Catalog_Pricing"
    Caption ="Item Search"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =3
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =90.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =2820
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =5400
                    Top =2520
                    Width =3900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="txMaterial"
                    Caption ="Fiber Type"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =2520
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =15480
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Fiber Cost"
                    LayoutCachedLeft =15480
                    LayoutCachedTop =2520
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =10980
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Fiber Count"
                    LayoutCachedLeft =10980
                    LayoutCachedTop =2520
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =2820
                End
                Begin Line
                    OverlapFlags =85
                    Left =240
                    Top =1140
                    Width =24240
                    Name ="Line58"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1140
                    LayoutCachedWidth =24480
                    LayoutCachedHeight =1140
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =8160
                    Top =300
                    Width =4740
                    Name ="Box35"
                    LayoutCachedLeft =8160
                    LayoutCachedTop =300
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =8280
                    Top =420
                    Width =4500
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Pricing Module"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =8280
                    LayoutCachedTop =420
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =915
                End
                Begin Label
                    OverlapFlags =85
                    Left =2040
                    Top =2520
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="txMake_Item"
                    Caption ="Price ($/m)"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =2520
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =2520
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label83"
                    Caption ="Base"
                    LayoutCachedLeft =360
                    LayoutCachedTop =2520
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =13980
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label85"
                    Caption ="Contribution"
                    LayoutCachedLeft =13980
                    LayoutCachedTop =2520
                    LayoutCachedWidth =15420
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =12480
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label87"
                    Caption ="Total Fiber Cost"
                    LayoutCachedLeft =12480
                    LayoutCachedTop =2520
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =9360
                    Top =2520
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label90"
                    Caption ="Multiplier"
                    LayoutCachedLeft =9360
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =2820
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =300
                    Top =1980
                    Width =3180
                    Height =285
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbProductCategory"
                    RowSourceType ="Table/Query"
                    RowSource ="dbo_tblSalesForce_Pricing"
                    AfterUpdate ="[Event Procedure]"
                    Format ="@;\" <search by category>\""

                    LayoutCachedLeft =300
                    LayoutCachedTop =1980
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =2265
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =300
                            Top =1680
                            Width =1545
                            Height =300
                            FontSize =10
                            FontWeight =500
                            Name ="Label96"
                            Caption ="Product Category"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1680
                            LayoutCachedWidth =1845
                            LayoutCachedHeight =1980
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3660
                    Top =1980
                    Width =4695
                    Height =285
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"140\""
                    Name ="cbFamilyDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblDesignTypes.SalesForceDescription, tblDesignTypes.Product_Pricing_Grou"
                        "p, tblDesignTypes.Active FROM tblDesignTypes WHERE (((tblDesignTypes.Product_Pri"
                        "cing_Group)=Forms!frmSalesForcePricing!cbProductCategory) And ((tblDesignTypes.A"
                        "ctive)<>0)); "
                    AfterUpdate ="[Event Procedure]"
                    Format ="@;\" <searc by family>\""

                    LayoutCachedLeft =3660
                    LayoutCachedTop =1980
                    LayoutCachedWidth =8355
                    LayoutCachedHeight =2265
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3660
                            Top =1680
                            Width =3780
                            Height =300
                            FontSize =10
                            FontWeight =500
                            Name ="Label95"
                            Caption ="Family Desc"
                            LayoutCachedLeft =3660
                            LayoutCachedTop =1680
                            LayoutCachedWidth =7440
                            LayoutCachedHeight =1980
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =16980
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label92"
                    Caption ="Base Cost"
                    LayoutCachedLeft =16980
                    LayoutCachedTop =2520
                    LayoutCachedWidth =18420
                    LayoutCachedHeight =2820
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =8460
                    Top =1980
                    Width =2760
                    Height =285
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbBase"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblCableConstructionReferences.Base, tblCableConstructionReferen"
                        "ces.ApprovedForQuoting, tblCableConstructionReferences.Active, tblDesignTypes.Sa"
                        "lesForceDescription FROM (tblCableConstructionReferences INNER JOIN tblDesignTyp"
                        "es ON tblCableConstructionReferences.DesignTypeID = tblDesignTypes.DesignTypeID)"
                        " INNER JOIN dbo_tblSalesForce_Catalog ON tblCableConstructionReferences.Base = d"
                        "bo_tblSalesForce_Catalog.Base WHERE (((tblCableConstructionReferences.ApprovedFo"
                        "rQuoting)<>0) And ((tblCableConstructionReferences.Active)<>0) And ((tblDesignTy"
                        "pes.SalesForceDescription) Like Forms!frmSalesForcePricing!cbFamilyDesc & \"*\")"
                        "); "
                    AfterUpdate ="[Event Procedure]"
                    Format ="@;\" <search by base>\""

                    LayoutCachedLeft =8460
                    LayoutCachedTop =1980
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =2265
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8460
                            Top =1680
                            Width =1845
                            Height =300
                            FontSize =10
                            FontWeight =500
                            Name ="Label97"
                            Caption ="Base"
                            LayoutCachedLeft =8460
                            LayoutCachedTop =1680
                            LayoutCachedWidth =10305
                            LayoutCachedHeight =1980
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =12000
                    Top =1920
                    Width =1320
                    Name ="btnSubmit"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12000
                    LayoutCachedTop =1920
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =2280
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =18480
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label99"
                    Caption ="FiberID"
                    LayoutCachedLeft =18480
                    LayoutCachedTop =2520
                    LayoutCachedWidth =19920
                    LayoutCachedHeight =2820
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15300
                    Top =1920
                    Width =1320
                    TabIndex =1
                    Name ="btnUpdatePricing"
                    Caption ="Update Pricing"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =15300
                    LayoutCachedTop =1920
                    LayoutCachedWidth =16620
                    LayoutCachedHeight =2280
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =3720
                    Top =2520
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label102"
                    Caption ="Price ($/ft)"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =2520
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =19980
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label104"
                    Caption ="Armored ($/m)"
                    LayoutCachedLeft =19980
                    LayoutCachedTop =2520
                    LayoutCachedWidth =21420
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =21480
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label105"
                    Caption ="Armor Adder"
                    LayoutCachedLeft =21480
                    LayoutCachedTop =2520
                    LayoutCachedWidth =22920
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =22980
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label106"
                    Caption ="Armor Size"
                    LayoutCachedLeft =22980
                    LayoutCachedTop =2520
                    LayoutCachedWidth =24420
                    LayoutCachedHeight =2820
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11280
                    Top =1920
                    Width =540
                    TabIndex =5
                    Name ="btnClearFields"
                    Caption ="Submit"
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

                    LayoutCachedLeft =11280
                    LayoutCachedTop =1920
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =2280
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
                    Left =22980
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="txArmor"
                    ControlSource ="Armor"
                    Format ="General Number"

                    LayoutCachedLeft =22980
                    LayoutCachedTop =60
                    LayoutCachedWidth =24420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21480
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =12
                    BorderColor =1643706
                    Name ="txArmorCostAdder"
                    ControlSource ="ArmorCostAdder"
                    Format ="General Number"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =21480
                    LayoutCachedTop =60
                    LayoutCachedWidth =22920
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19980
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    BorderColor =-2
                    Name ="txArmoredPrice"
                    ControlSource ="ArmoredPrice"

                    LayoutCachedLeft =19980
                    LayoutCachedTop =60
                    LayoutCachedWidth =21420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3720
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="txPrice_ft"
                    ControlSource ="=Round([Price]/3.21,4)"
                    Format ="General Number"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18480
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="Text100"
                    ControlSource ="FiberID"
                    Format ="General Number"

                    LayoutCachedLeft =18480
                    LayoutCachedTop =60
                    LayoutCachedWidth =19920
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16980
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="txBaseCost"
                    ControlSource ="BaseSkeletalCost"
                    Format ="General Number"

                    LayoutCachedLeft =16980
                    LayoutCachedTop =60
                    LayoutCachedWidth =18420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9360
                    Top =60
                    Width =1560
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    BorderColor =1643706
                    Name ="txMultiplier"
                    ControlSource ="Multiplier"
                    Format ="General Number"
                    AfterUpdate ="[Event Procedure]"
                    OnDirty ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =60
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12480
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="txTotalFiberCost"
                    ControlSource ="TotalFiberCost"
                    Format ="General Number"

                    LayoutCachedLeft =12480
                    LayoutCachedTop =60
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="txBase"
                    ControlSource ="Base"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13980
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    BorderColor =1643706
                    Name ="txContribution"
                    ControlSource ="Contribution"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =13980
                    LayoutCachedTop =60
                    LayoutCachedWidth =15420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txPrice"
                    ControlSource ="Price"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10980
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="txFiberCount"
                    ControlSource ="FiberCount"
                    Format ="General Number"

                    LayoutCachedLeft =10980
                    LayoutCachedTop =60
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15480
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="txCPQ_Fiber_Cost"
                    ControlSource ="CPQ_Fiber_Cost"

                    LayoutCachedLeft =15480
                    LayoutCachedTop =60
                    LayoutCachedWidth =16920
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5400
                    Top =60
                    Width =3900
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txFiberType"
                    ControlSource ="Fiber_Type__c"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =900
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =300
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =300
                    LayoutCachedTop =300
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =660
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1860
                    Top =300
                    Width =540
                    TabIndex =1
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

                    LayoutCachedLeft =1860
                    LayoutCachedTop =300
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =660
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
Dim oldMultiplier As Double


Private Sub txFinishedGood_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

Private Sub btnClearFields_Click()
Me.cbProductCategory = Null
Me.cbProductCategory.Requery
cbProductCategory_AfterUpdate
Form_Load
End Sub

Private Sub btnClose_Click()
DoCmd.Close acForm, Me.name
End Sub


Private Sub btnExcelExport_Click()
On Error GoTo ErrorHandler

DoCmd.OutputTo acOutputForm, "frmBOMCostExplode", acFormatXLS, Environ("USERPROFILE") & "\Desktop\SalesForcePricing.xls"
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

Private Sub btnSubmit_Click()
If IsNull(Me.cbBase) Then
    Exit Sub
End If

GetSalesForcePricing (Me.cbBase)
Me.Requery
End Sub

Private Sub btnUpdatePricing_Click()
DoCmd.SetWarnings False

Me.btnUpdatePricing.SetFocus
Me.Refresh
DoCmd.OpenQuery "qrySalesForceUpdatePricingMultiplier"
DoCmd.OpenQuery "qrySalesForceUpdatePricingContribution"
DoCmd.OpenQuery "qrySalesForceUpdatePricingArmorCostAdder"
Call btnSubmit_Click

End Sub

Private Sub cbBase_AfterUpdate()
btnSubmit_Click
End Sub

Private Sub cbFamilyDesc_AfterUpdate()
Me.cbBase = Null
Me.cbBase.Requery
End Sub

Private Sub cbProductCategory_AfterUpdate()
Me.cbFamilyDesc = Null
Me.cbBase = Null
Me.cbFamilyDesc.Requery
Me.cbBase.Requery

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub


Private Sub Form_Load()
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE FROM tblSalesForce_Catalog_Pricing"
Me.Requery
End Sub



Private Sub txArmorCostAdder_AfterUpdate()
DoCmd.SetWarnings False
Me.Refresh
DoCmd.RunSQL "UPDATE tblSalesForce_Catalog_Pricing SET ArmorCostAdder = " & Me.txArmorCostAdder
End Sub

Private Sub txContribution_AfterUpdate()
DoCmd.SetWarnings False
Me.Refresh
DoCmd.RunSQL "UPDATE tblSalesForce_Catalog_Pricing SET Contribution = " & Me.Contribution
End Sub

Private Sub txMultiplier_AfterUpdate()
If Me.Multiplier >= 1 Then
    Me.Multiplier = oldMultiplier
End If

End Sub

Private Sub txMultiplier_Dirty(Cancel As Integer)
oldMultiplier = Me.Multiplier
End Sub

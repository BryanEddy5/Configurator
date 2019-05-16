Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =22
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =24375
    DatasheetFontHeight =10
    ItemSuffix =213
    Left =1650
    Top =480
    Right =22980
    Bottom =13455
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa6fef19fd09be440
    End
    Caption ="Create New Item"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin FormHeader
            CanGrow = NotDefault
            Height =540
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =9000
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =5760
                    Left =8460
                    Top =1200
                    Width =3420
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Customer"
                    RowSourceType ="Table/Query"
                    RowSource ="vCustomers"
                    ColumnWidths ="1440;0;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =8460
                    Top =1620
                    Width =3420
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="PID"
                    RowSourceType ="Table/Query"
                    RowSource ="qryCreateItemPID"
                    ColumnWidths ="2160;0;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    Format =">"

                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    DecimalPlaces =5
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =8640
                    Left =2100
                    Top =3300
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =18
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Fiber1"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemFiberType"
                    ColumnWidths ="2160;864;1440;1440;1440;0"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    FontName ="Times New Roman"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =3300
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =3600
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =780
                    Width =705
                    Height =300
                    Name ="Label48"
                    Caption ="Oracle #"
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =1200
                    Width =705
                    Height =300
                    Name ="Label49"
                    Caption ="Catalog"
                End
                Begin Label
                    OverlapFlags =85
                    Left =6900
                    Top =1200
                    Width =840
                    Height =300
                    Name ="Label50"
                    Caption ="Customer"
                End
                Begin Label
                    OverlapFlags =85
                    Left =6900
                    Top =1620
                    Width =705
                    Height =300
                    Name ="Label51"
                    Caption ="PID"
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =2460
                    Width =705
                    Height =300
                    Name ="Label52"
                    Caption ="Jacket"
                    LayoutCachedLeft =780
                    LayoutCachedTop =2460
                    LayoutCachedWidth =1485
                    LayoutCachedHeight =2760
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =3300
                    Width =705
                    Height =300
                    Name ="Label55"
                    Caption ="Fiber"
                    LayoutCachedLeft =780
                    LayoutCachedTop =3300
                    LayoutCachedWidth =1485
                    LayoutCachedHeight =3600
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =8640
                    Left =8460
                    Top =4140
                    Width =7740
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"140\""
                    Name ="PrintLine1"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemPrintLine1"
                    ColumnWidths ="7920;0;0"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =4140
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =4440
                End
                Begin Label
                    OverlapFlags =85
                    Left =6900
                    Top =4140
                    Width =840
                    Height =240
                    Name ="Label56"
                    Caption ="Print Line 1"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =17880
                    Top =6060
                    Width =780
                    Height =300
                    Name ="Label53"
                    Caption ="ColorChip"
                    LayoutCachedLeft =17880
                    LayoutCachedTop =6060
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =6360
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =2880
                    Width =705
                    Height =300
                    Name ="Label54"
                    Caption ="Color"
                    LayoutCachedLeft =780
                    LayoutCachedTop =2880
                    LayoutCachedWidth =1485
                    LayoutCachedHeight =3180
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2100
                    Top =2880
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"30\""
                    Name ="Color"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemChipColor"
                    ColumnWidths ="2880;0;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    Format =">"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =2880
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =3180
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =8460
                    Top =780
                    Width =3420
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="PrintType"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmFiberTypeStandard"
                    ColumnWidths ="0;2160;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    OnGotFocus ="[Event Procedure]"
                    Format =">"

                End
                Begin Label
                    OverlapFlags =85
                    Left =6900
                    Top =720
                    Width =840
                    Height =360
                    Name ="Label1"
                    Caption ="Print Type"
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =3720
                    Width =705
                    Height =300
                    Name ="Label3"
                    Caption ="OM#"
                    LayoutCachedLeft =780
                    LayoutCachedTop =3720
                    LayoutCachedWidth =1485
                    LayoutCachedHeight =4020
                End
                Begin Label
                    OverlapFlags =85
                    Left =4380
                    Top =3720
                    Width =360
                    Height =300
                    Name ="Label5"
                    Caption ="BIF"
                    LayoutCachedLeft =4380
                    LayoutCachedTop =3720
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =4020
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =8640
                    Left =8460
                    Top =4560
                    Width =7740
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"140\""
                    Name ="PrintLine2"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreatItemPrintLine2"
                    ColumnWidths ="7920;432;0"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =4560
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =4860
                End
                Begin Label
                    OverlapFlags =85
                    Left =6900
                    Top =4560
                    Width =840
                    Height =240
                    Name ="Label7"
                    Caption ="Print Line 2"
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =8640
                    Left =8460
                    Top =4980
                    Width =7740
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="PrintLine3"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemPrintLine3"
                    ColumnWidths ="7920;432;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =4980
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =5280
                End
                Begin Label
                    OverlapFlags =85
                    Left =6900
                    Top =4980
                    Width =870
                    Height =240
                    Name ="Label9"
                    Caption ="Print Line 3"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5040
                    Left =8460
                    Top =5400
                    Width =7740
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="PrintLine4"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemPrintLine4"
                    ColumnWidths ="4320;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =5400
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =5685
                End
                Begin Label
                    OverlapFlags =85
                    Left =6900
                    Top =5400
                    Width =840
                    Height =240
                    Name ="Label11"
                    Caption ="Print Line 4"
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =8460
                    Top =2880
                    Width =2475
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="ListingCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemListing"
                    ColumnWidths ="1152;0;0;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    OnGotFocus ="[Event Procedure]"
                    Format ="@;\" <pick listing company>\""

                    LayoutCachedLeft =8460
                    LayoutCachedTop =2880
                    LayoutCachedWidth =10935
                    LayoutCachedHeight =3180
                End
                Begin Label
                    OverlapFlags =85
                    Left =6900
                    Top =2880
                    Width =1155
                    Height =240
                    Name ="label12"
                    Caption ="Listing Comany"
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =1620
                    Width =705
                    Height =300
                    Name ="Label17"
                    Caption ="Base"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5820
                    Top =3300
                    Width =480
                    TabIndex =23
                    Name ="FiberSpecs"
                    Caption ="Spec"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =3300
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =3660
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =8
                    ListWidth =11520
                    Left =2100
                    Top =1620
                    Width =3660
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";"
                        "\"10\";\"100\""
                    Name ="Base"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * FROM qryfrmCreateItemBaseNew WHERE DesignExtension = ':CCC'"
                    ColumnWidths ="2160;2880;720;1440;1440;1440;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1620
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1905
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =19440
                    Top =5700
                    TabIndex =24
                    Name ="Unprinted"
                    DefaultValue ="-1"

                    LayoutCachedLeft =19440
                    LayoutCachedTop =5700
                    LayoutCachedWidth =19700
                    LayoutCachedHeight =5940
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =19800
                            Top =5700
                            Width =1620
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="text1"
                            Caption ="Unprinted"
                            LayoutCachedLeft =19800
                            LayoutCachedTop =5700
                            LayoutCachedWidth =21420
                            LayoutCachedHeight =5940
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =1080
                    Width =780
                    TabIndex =27
                    BackColor =26367
                    Name ="Location"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =1080
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =1320
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4140
                    Top =5880
                    Width =1320
                    Height =420
                    TabIndex =25
                    Name ="btnUnitSeries"
                    Caption ="Unit Series"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =780
                    Width =780
                    TabIndex =28
                    BackColor =26367
                    Name ="Exists"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =780
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =1020
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4140
                    Top =6360
                    Width =1320
                    Height =405
                    TabIndex =26
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
                    Left =14280
                    Top =3360
                    Width =600
                    Height =600
                    TabIndex =29
                    Name ="btnPrint"
                    Caption ="Print"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =14280
                    LayoutCachedTop =3360
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =3960
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =1380
                    Width =780
                    TabIndex =30
                    BackColor =26367
                    Name ="Armor"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =1380
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =1620
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =1680
                    Width =780
                    TabIndex =33
                    BackColor =26367
                    Name ="ArmorType"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =1680
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =1920
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =1980
                    Width =780
                    TabIndex =34
                    BackColor =26367
                    Name ="ArmorNo"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =1980
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =2220
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =20400
                    Top =780
                    Width =840
                    Height =240
                    Name ="Label69"
                    Caption ="Exists"
                    LayoutCachedLeft =20400
                    LayoutCachedTop =780
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =1020
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =20400
                    Top =1080
                    Width =840
                    Height =240
                    Name ="Label70"
                    Caption ="Location"
                    LayoutCachedLeft =20400
                    LayoutCachedTop =1080
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =1320
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =20400
                    Top =1380
                    Width =840
                    Height =240
                    Name ="Label71"
                    Caption ="Armor"
                    LayoutCachedLeft =20400
                    LayoutCachedTop =1380
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =1620
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =20400
                    Top =1680
                    Width =870
                    Height =240
                    Name ="Label72"
                    Caption ="ArmorType"
                    LayoutCachedLeft =20400
                    LayoutCachedTop =1680
                    LayoutCachedWidth =21270
                    LayoutCachedHeight =1920
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =20400
                    Top =1980
                    Width =840
                    Height =240
                    Name ="Label73"
                    Caption ="ArmorNo"
                    LayoutCachedLeft =20400
                    LayoutCachedTop =1980
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =2220
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2100
                    Top =6000
                    Width =1080
                    Height =360
                    FontSize =11
                    FontWeight =600
                    TabIndex =35
                    Name ="NomOD"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =6000
                            Width =1125
                            Height =240
                            Name ="Label75"
                            Caption ="Nom OD (mm):"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =2280
                    Width =780
                    TabIndex =36
                    BackColor =26367
                    Name ="ArmorCoreOD"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =2280
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =2520
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =20400
                    Top =2280
                    Width =840
                    Height =240
                    Name ="Label77"
                    Caption ="ArmorOD"
                    LayoutCachedLeft =20400
                    LayoutCachedTop =2280
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =2520
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =1080
                    Width =780
                    TabIndex =37
                    BackColor =26367
                    Name ="SubColor"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =1080
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =1320
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =22980
                    Top =1080
                    Width =840
                    Height =240
                    Name ="Label79"
                    Caption ="SubColor"
                    LayoutCachedLeft =22980
                    LayoutCachedTop =1080
                    LayoutCachedWidth =23820
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =1680
                    Width =780
                    TabIndex =38
                    BackColor =26367
                    Name ="FiberCount"
                    ControlSource ="=CDbl(Mid(Forms!frmCreateItem!ItemNo,3,3))"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =1680
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =1920
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =22980
                    Top =1680
                    Width =870
                    Height =240
                    Name ="Label81"
                    Caption ="FiberCount"
                    LayoutCachedLeft =22980
                    LayoutCachedTop =1680
                    LayoutCachedWidth =23850
                    LayoutCachedHeight =1920
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =1980
                    Width =780
                    TabIndex =39
                    BackColor =26367
                    Name ="SubCount"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =1980
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =2220
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =22980
                    Top =1980
                    Width =870
                    Height =240
                    Name ="Label85"
                    Caption ="SubCount"
                    LayoutCachedLeft =22980
                    LayoutCachedTop =1980
                    LayoutCachedWidth =23850
                    LayoutCachedHeight =2220
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =2280
                    Width =780
                    TabIndex =40
                    BackColor =26367
                    Name ="SubNo"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =2280
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =2520
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =22980
                    Top =2280
                    Width =870
                    Height =240
                    Name ="Label87"
                    Caption ="SubNo"
                    LayoutCachedLeft =22980
                    LayoutCachedTop =2280
                    LayoutCachedWidth =23850
                    LayoutCachedHeight =2520
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =1380
                    Width =780
                    TabIndex =41
                    BackColor =26367
                    Name ="SubColorCount"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =1380
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =1620
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =22980
                    Top =1380
                    Width =1155
                    Height =240
                    Name ="Label97"
                    Caption ="SubColorCount"
                    LayoutCachedLeft =22980
                    LayoutCachedTop =1380
                    LayoutCachedWidth =24135
                    LayoutCachedHeight =1620
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =2580
                    Width =780
                    TabIndex =42
                    BackColor =26367
                    Name ="MixedSub"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =2580
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =2820
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =22980
                    Top =2580
                    Width =870
                    Height =240
                    Name ="Label99"
                    Caption ="MixedSub"
                    LayoutCachedLeft =22980
                    LayoutCachedTop =2580
                    LayoutCachedWidth =23850
                    LayoutCachedHeight =2820
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =8640
                    Left =2100
                    Top =4260
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Fiber2"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemFiber2Type"
                    ColumnWidths ="2160;864;1440;1440;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    Format =">"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =4260
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =4560
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =4260
                            Width =705
                            Height =300
                            Name ="Label101"
                            Caption ="Fiber2"
                            LayoutCachedLeft =780
                            LayoutCachedTop =4260
                            LayoutCachedWidth =1485
                            LayoutCachedHeight =4560
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =2940
                    Width =780
                    TabIndex =45
                    BackColor =26367
                    Name ="MixedFiber"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =2940
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =3180
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =2940
                            Width =870
                            Height =240
                            Name ="Label107"
                            Caption ="MixedFiber"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =2940
                            LayoutCachedWidth =21270
                            LayoutCachedHeight =3180
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =3240
                    Width =780
                    TabIndex =46
                    BackColor =26367
                    Name ="Fiber12"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =3240
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =3480
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =3240
                            Width =870
                            Height =240
                            Name ="Label109"
                            Caption ="Fiber1/2"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =3240
                            LayoutCachedWidth =21270
                            LayoutCachedHeight =3480
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =3540
                    Width =780
                    TabIndex =47
                    BackColor =26367
                    Name ="Fiber1Count"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =3540
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =3780
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =3540
                            Width =1260
                            Height =240
                            Name ="Label115"
                            Caption ="Fiber1Count"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =3540
                            LayoutCachedWidth =21660
                            LayoutCachedHeight =3780
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =3840
                    Width =780
                    TabIndex =48
                    BackColor =26367
                    Name ="Fiber2Count"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =3840
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =4080
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =3840
                            Width =1200
                            Height =240
                            Name ="Label117"
                            Caption ="Fiber2Count"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =3840
                            LayoutCachedWidth =21600
                            LayoutCachedHeight =4080
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =4140
                    Width =780
                    TabIndex =49
                    BackColor =26367
                    Name ="Fiber1Type"
                    ControlSource ="=IIf([ItemNo] Like \"*/*\",Mid([ItemNo],InStr(12,[ItemNo],\"MM\")+2,1),Mid([Item"
                        "No],6,1))"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =4140
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =4380
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =4140
                            Width =1260
                            Height =240
                            Name ="Label121"
                            Caption ="Fiber1Type"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =4140
                            LayoutCachedWidth =21660
                            LayoutCachedHeight =4380
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =4440
                    Width =780
                    TabIndex =50
                    BackColor =26367
                    Name ="Fiber2Type"
                    ControlSource ="=IIf([ItemNo] Like \"*/*\",Mid([ItemNo],InStr(12,[ItemNo],\"SM\")+2,1),\"\")"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =4440
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =4680
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =4440
                            Width =1200
                            Height =240
                            Name ="Label123"
                            Caption ="Fiber2Type"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =4440
                            LayoutCachedWidth =21600
                            LayoutCachedHeight =4680
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =4140
                    Width =780
                    TabIndex =51
                    BackColor =26367
                    Name ="Fiber"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =4140
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =4380
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =4140
                            Width =870
                            Height =240
                            Name ="Label125"
                            Caption ="Fiber"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =4140
                            LayoutCachedWidth =23850
                            LayoutCachedHeight =4380
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =4440
                    Width =780
                    TabIndex =52
                    BackColor =26367
                    Name ="Sub1Count"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =4440
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =4680
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =4440
                            Width =1260
                            Height =240
                            Name ="Label127"
                            Caption ="Sub1Count"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =4440
                            LayoutCachedWidth =24240
                            LayoutCachedHeight =4680
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =4740
                    Width =780
                    TabIndex =53
                    BackColor =26367
                    Name ="Sub2Count"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =4740
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =4980
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =4740
                            Width =1200
                            Height =240
                            Name ="Label129"
                            Caption ="Sub2Count"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =4740
                            LayoutCachedWidth =24180
                            LayoutCachedHeight =4980
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =3840
                    Width =780
                    TabIndex =54
                    BackColor =26367
                    Name ="FiberType"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =3840
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =4080
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =3840
                            Width =870
                            Height =240
                            Name ="Label131"
                            Caption ="FiberType"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =3840
                            LayoutCachedWidth =23850
                            LayoutCachedHeight =4080
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =780
                    Width =780
                    TabIndex =55
                    BackColor =26367
                    Name ="TBColor"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =780
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =780
                            Width =840
                            Height =240
                            Name ="Label135"
                            Caption ="TBColor"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =780
                            LayoutCachedWidth =23820
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =5040
                    Width =780
                    TabIndex =56
                    BackColor =26367
                    Name ="MixedTB"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =5040
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =5280
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =5040
                            Width =1200
                            Height =240
                            Name ="Label137"
                            Caption ="MixedTB"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =5040
                            LayoutCachedWidth =24180
                            LayoutCachedHeight =5280
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =5340
                    Width =780
                    TabIndex =57
                    BackColor =26367
                    Name ="StripePrint"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =5340
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =5580
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =5340
                            Width =1200
                            Height =240
                            Name ="Label139"
                            Caption ="StripePrint"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =5340
                            LayoutCachedWidth =24180
                            LayoutCachedHeight =5580
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =4740
                    Width =780
                    TabIndex =58
                    BackColor =26367
                    Name ="SubX"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =4740
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =4980
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =4740
                            Width =1200
                            Height =240
                            Name ="Label141"
                            Caption ="SubX"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =4740
                            LayoutCachedWidth =21600
                            LayoutCachedHeight =4980
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12720
                    Top =3780
                    Width =1260
                    TabIndex =59
                    Name ="ListingNumber"

                    LayoutCachedLeft =12720
                    LayoutCachedTop =3780
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =4020
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11160
                            Top =3780
                            Width =1200
                            Height =240
                            Name ="Label143"
                            Caption ="E Number"
                            LayoutCachedLeft =11160
                            LayoutCachedTop =3780
                            LayoutCachedWidth =12360
                            LayoutCachedHeight =4020
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2100
                    Top =780
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NewOracle#"
                    Format =">"
                    FontName ="Times New Roman"
                    EventProcPrefix ="NewOracle_"

                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2100
                    Top =1200
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="ItemNo"
                    Format =">"
                    AfterUpdate ="mcrReQueryItemNo"
                    DefaultValue ="=Trim(UCase([Forms]![frmMainMenu]![ItemNo]))"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =3540
                    Width =780
                    TabIndex =60
                    BackColor =26367
                    Name ="FiberType2"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =3540
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =3780
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =3540
                            Width =885
                            Height =240
                            Name ="Label145"
                            Caption ="FiberType2"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =3540
                            LayoutCachedWidth =23865
                            LayoutCachedHeight =3780
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =5820
                    Top =4260
                    Width =480
                    TabIndex =61
                    Name ="FiberSpecs2"
                    Caption ="Spec"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =4260
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =4620
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =3240
                    Width =780
                    TabIndex =62
                    BackColor =26367
                    Name ="TBY"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =3240
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =3480
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =3240
                            Width =885
                            Height =240
                            Name ="Label148"
                            Caption ="TBY"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =3240
                            LayoutCachedWidth =23865
                            LayoutCachedHeight =3480
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2100
                    Top =2460
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =15
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="Jacket"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemJacketType"
                    ColumnWidths ="3600;0;0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="Null"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =2460
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =2580
                    Width =780
                    TabIndex =63
                    BackColor =26367
                    Name ="JacketColor"
                    ControlSource ="=Mid([ItemNo],10,1)"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =2580
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =2820
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =2580
                            Width =885
                            Height =240
                            Name ="Label152"
                            Caption ="JacketColor"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =2580
                            LayoutCachedWidth =21285
                            LayoutCachedHeight =2820
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2100
                    Top =3720
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =31
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="OM"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =3720
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =4020
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4980
                    Top =3720
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =32
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="BIF"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =3720
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =4020
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4980
                    Top =4680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =44
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="BIF2"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =4680
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =4980
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4380
                            Top =4680
                            Width =360
                            Height =300
                            Name ="Label103"
                            Caption ="BIF"
                            LayoutCachedLeft =4380
                            LayoutCachedTop =4680
                            LayoutCachedWidth =4740
                            LayoutCachedHeight =4980
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2100
                    Top =4680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =43
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="OM2"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =4680
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =4980
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =4680
                            Width =705
                            Height =300
                            Name ="Label102"
                            Caption ="OM#"
                            LayoutCachedLeft =780
                            LayoutCachedTop =4680
                            LayoutCachedWidth =1485
                            LayoutCachedHeight =4980
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =2940
                    Width =780
                    TabIndex =64
                    BackColor =26367
                    Name ="FiberExt"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =2940
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =3180
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =2940
                            Width =885
                            Height =240
                            Name ="Label154"
                            Caption ="FiberExt"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =2940
                            LayoutCachedWidth =23865
                            LayoutCachedHeight =3180
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =8460
                    Top =3300
                    Width =2475
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="ListingType"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemBurnListingUS"
                    ColumnWidths ="1152;0;0;0;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    Format ="@;\" <pick rating>\""

                    LayoutCachedLeft =8460
                    LayoutCachedTop =3300
                    LayoutCachedWidth =10935
                    LayoutCachedHeight =3585
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6900
                            Top =3300
                            Width =1080
                            Height =240
                            Name ="Label156"
                            Caption ="Listing US"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8460
                    Top =3720
                    Width =2475
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="ListingTypeCanada"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemBurnListingCanada"
                    ColumnWidths ="1152;432;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    Format =">"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =3720
                    LayoutCachedWidth =10935
                    LayoutCachedHeight =4020
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6900
                            Top =3720
                            Width =1125
                            Height =240
                            Name ="Label159"
                            Caption ="Listing Canada"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =5040
                    Width =780
                    TabIndex =65
                    BackColor =26367
                    Name ="TBType"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =5040
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =5280
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =5040
                            Width =1200
                            Height =240
                            Name ="Label163"
                            Caption ="TBType"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =5040
                            LayoutCachedWidth =21600
                            LayoutCachedHeight =5280
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12720
                    Top =2880
                    Width =1260
                    TabIndex =20
                    Name ="ListingConstruction"

                    LayoutCachedLeft =12720
                    LayoutCachedTop =2880
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =3120
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11160
                            Top =2880
                            Width =1440
                            Height =240
                            Name ="Label88888"
                            Caption ="ListingConstruction"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12720
                    Top =3180
                    Width =1260
                    TabIndex =21
                    Name ="ListingSection"

                    LayoutCachedLeft =12720
                    LayoutCachedTop =3180
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =3420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11160
                            Top =3180
                            Width =1200
                            Height =240
                            Name ="Label174"
                            Caption ="ListingSection"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12720
                    Top =3480
                    Width =1260
                    TabIndex =22
                    Name ="Listing"

                    LayoutCachedLeft =12720
                    LayoutCachedTop =3480
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =3720
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11160
                            Top =3480
                            Width =1200
                            Height =240
                            Name ="Label176"
                            Caption ="Listing"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21960
                    Top =480
                    Width =780
                    TabIndex =66
                    BackColor =26367
                    Name ="HybridSubColor"

                    LayoutCachedLeft =21960
                    LayoutCachedTop =480
                    LayoutCachedWidth =22740
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =22980
                            Top =480
                            Width =1395
                            Height =240
                            Name ="Label178"
                            Caption ="HybridSubColor"
                            LayoutCachedLeft =22980
                            LayoutCachedTop =480
                            LayoutCachedWidth =24375
                            LayoutCachedHeight =720
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =19200
                    Top =6060
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =17
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="ColorChip"
                    Format =">"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =19200
                    LayoutCachedTop =6060
                    LayoutCachedWidth =22860
                    LayoutCachedHeight =6360
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8460
                    Top =2460
                    Width =1980
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="CustomerRev"
                    Format ="@;\" <enter spec rev>\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =2460
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =2760
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6900
                            Top =2460
                            Width =705
                            Height =300
                            Name ="Label180"
                            Caption ="Rev"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3120
                    Top =4680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =67
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Mode2"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3120
                    LayoutCachedTop =4680
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =4980
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3120
                    Top =3720
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =68
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Mode"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3120
                    LayoutCachedTop =3720
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =4020
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19380
                    Top =5340
                    Width =780
                    TabIndex =69
                    BackColor =26367
                    Name ="PrintWhite"

                    LayoutCachedLeft =19380
                    LayoutCachedTop =5340
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =5580
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20400
                            Top =5340
                            Width =1200
                            Height =240
                            Name ="Label185"
                            Caption ="PrintWhite"
                            LayoutCachedLeft =20400
                            LayoutCachedTop =5340
                            LayoutCachedWidth =21600
                            LayoutCachedHeight =5580
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8460
                    Top =5850
                    Width =420
                    Height =300
                    TabIndex =70
                    Name ="StripedPrint"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6900
                            Top =5820
                            Width =1050
                            Height =240
                            Name ="Label187"
                            Caption ="Striped Print"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12240
                    Top =1020
                    Width =1086
                    Height =606
                    TabIndex =71
                    Name ="btnAttachSpec"
                    Caption ="Attach Spec"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Find Next"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =18120
                    Top =780
                    Width =1086
                    Height =606
                    TabIndex =72
                    Name ="btnSaveSpec"
                    Caption ="Save Spec"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"

                    LayoutCachedLeft =18120
                    LayoutCachedTop =780
                    LayoutCachedWidth =19206
                    LayoutCachedHeight =1386
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8460
                    Top =2040
                    Width =1980
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Spec"
                    Format ="@;\" <enter spec name>\""
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =2040
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =2340
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6900
                            Top =2040
                            Width =1080
                            Height =240
                            Name ="Label195"
                            Caption ="Spec Doc #"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =12240
                    Top =1740
                    Width =1086
                    Height =606
                    TabIndex =73
                    Name ="btnShowSpec"
                    Caption ="Grab Spec"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =18120
                    Top =1500
                    Width =1086
                    Height =606
                    TabIndex =74
                    Name ="btnOpenSpec"
                    Caption ="Show Spec"
                    ControlTipText ="Find Next"

                    LayoutCachedLeft =18120
                    LayoutCachedTop =1500
                    LayoutCachedWidth =19206
                    LayoutCachedHeight =2106
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19080
                    Top =420
                    Width =780
                    TabIndex =75
                    BackColor =26367
                    Name ="strTB"

                    LayoutCachedLeft =19080
                    LayoutCachedTop =420
                    LayoutCachedWidth =19860
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =20100
                            Top =420
                            Width =1695
                            Height =240
                            Name ="Label199"
                            Caption ="TB Creation : TB picked"
                            LayoutCachedLeft =20100
                            LayoutCachedTop =420
                            LayoutCachedWidth =21795
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5880
                    Top =1500
                    Width =426
                    Height =426
                    TabIndex =76
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
                        0xadada000000dadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =5880
                    LayoutCachedTop =1500
                    LayoutCachedWidth =6306
                    LayoutCachedHeight =1926
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =2040
                    Width =705
                    Height =300
                    Name ="Label206"
                    Caption ="Subunit"
                    LayoutCachedLeft =780
                    LayoutCachedTop =2040
                    LayoutCachedWidth =1485
                    LayoutCachedHeight =2340
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =7200
                    Left =2100
                    Top =2040
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =77
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbSubunit"
                    RowSourceType ="Table/Query"
                    RowSource ="qryfrmCreateItemSubunitSelect"
                    ColumnWidths ="2160;2880;720"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =2340
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5880
                    Top =2040
                    Width =426
                    Height =426
                    TabIndex =78
                    Name ="btnFindSub"
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
                        0xadada000000dadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =5880
                    LayoutCachedTop =2040
                    LayoutCachedWidth =6306
                    LayoutCachedHeight =2466
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13560
                    Top =900
                    Width =4035
                    Height =1425
                    TabIndex =79
                    Name ="txCustomerNotes"

                    LayoutCachedLeft =13560
                    LayoutCachedTop =900
                    LayoutCachedWidth =17595
                    LayoutCachedHeight =2325
                    Begin
                        Begin Label
                            FontUnderline = NotDefault
                            OverlapFlags =85
                            Left =13560
                            Top =480
                            Width =1560
                            Height =300
                            FontSize =10
                            Name ="Label210"
                            Caption ="Customer Notes:"
                            LayoutCachedLeft =13560
                            LayoutCachedTop =480
                            LayoutCachedWidth =15120
                            LayoutCachedHeight =780
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =2460
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4140
                    Top =180
                    Width =1320
                    Name ="btnSubmit"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =180
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =540
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1740
                    Top =240
                    Width =2040
                    Height =480
                    TabIndex =1
                    Name ="UnitSeries"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =360
                            Top =300
                            Width =840
                            Height =240
                            Name ="Label57"
                            Caption ="Unit Series"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4140
                    Top =660
                    Width =1320
                    TabIndex =2
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7680
                    Top =480
                    Width =5100
                    TabIndex =3
                    Name ="FileList"
                    RowSourceType ="Value List"

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
Dim RibbonHighCure As Boolean
Dim CableConstruction As New clsCableConstruction
Dim ItemCustomer As New clsCustomer
Dim bBronzeBraided  As Boolean
Private mlFiber1 As New clsFiber
Private mlFiber2 As New clsFiber
Private mlDesignCode As String
Private CableColorIndicator As String
Dim FillerTube As Boolean
Private WithEvents m_frmMainMenu As Form_frmMainMenu
Attribute m_frmMainMenu.VB_VarHelpID = -1
Private m_ItemRequestNumber

Friend Property Get frmMainMenu() As Form_frmMainMenu
    If m_frmMainMenu Is Nothing Then Set m_frmMainMenu = Form_frmMainMenu
    Set frmMainMenu = m_frmMainMenu
End Property

Friend Property Get ItemRequestNumber() As String

    ItemRequestNumber = m_ItemRequestNumber



End Property

Private Sub AfterBaseSelection()

    
    If Me.Armor <> -1 And Not IsNull(Me.Base) Then
    
        Me.Jacket.Requery
        Me.Jacket = Me.Jacket.ItemData(1)
        Me.Color.Requery
        Me.Color = Me.Color.ItemData(1)
        Me.NomOD = Me.Base.Column(2)
        Call Color_AfterUpdate
        
        CableConstruction.Load (Me.Base)
        Me.Location = CableConstruction.Location
        
        'Check for subunits if item is a two pass
        If IsNull(Me.cbSubunit.ItemData(2)) = False And CableConstruction.Location = 2 Then
            MsgBox ("Multiple Possible Subunits.  Review"), vbCritical
        Else
            Me.cbSubunit.Requery
            Me.cbSubunit = Me.cbSubunit.ItemData(1)
        End If
        
        
        'Get color chip
        If Jacket = "NONE" Then
            Me.Color = CreateNewItem.getJacketColor(CableColorIndicator)
    
        Else
            Me.Color.RowSourceType = "Table/Query"
        End If
    
        
        
        
        If bFurcation And Mid(Me.ItemNo, 11, 1) <> "0" Then
            If FindFurcation(Me.Base, Me.Location, Mid(Me.ItemNo, 11, 1)) = False Then
                Me.btnSubmit.Enabled = False
            Else
                Me.btnSubmit.Enabled = True
            End If
        End If
    Else
        Me.cbSubunit.Enabled = False
    
    End If

    If IsNull(Me.Customer) = False Then
        Call Customer_AfterUpdate
    End If
       
    'Identifies if there is a class 1 fiber and runs the sub routine
    If Mid(Me.ItemNo, 12, 1) = "B" And Mid(Me.ItemNo, 11, 1) = "P" Then
        Call Class1Fiber
        Call Fiber1_AfterUpdate
    Else
    
        Call Fiber1_GotFocus
        Me.Fiber1.Requery
        If Me.ItemNo Like "?????@*" Then         'Assign the fiber type of "NONE" for furcation cables
            Me.Fiber1 = "NONE"
        Else
            If IsNull(Me.Fiber1.ItemData(1)) Then 'Check if any fiber is available
                MsgBox "There is no fiber for this indicator and fiber configuration." & vbCrLf & "Check the design code and rebuild.", vbCritical
                Me.btnSubmit.Enabled = False
            Else
                Me.Fiber1 = Me.Fiber1.ItemData(1) 'Else pick the first fiber from the combobox
            End If
        End If
        Call Fiber1_AfterUpdate
    End If
    
    ' Identifies if there is more than one fiber and automatically picks the cheapest available fiber.
    If Me.MixedFiber = 1 Then
        Me.Fiber2.Requery
        Me.Fiber2 = Me.Fiber2.ItemData(1)
        Call Fiber2_AfterUpdate
    End If
    
    Call FiberPrint


    If CableConstruction.Location = 1 Or Forms!frmCreateItem!Armor = -1 Then
        btnUnitSeries.Visible = False
        Me.UnitSeries.Visible = False
        Me.cbSubunit.Enabled = False

    Else
        btnUnitSeries.Visible = True
        Me.UnitSeries.Visible = True
        Me.cbSubunit.Enabled = True
        DoCmd.Close acForm, "frmCreateItemUnitSeriesExist"
    End If

    Call CableAttributes
    
    'Check to see if a 0 color indicator has been requested.
    Call ItemBuildChecks
    
End Sub

Public Sub Autofill()
    'Dim rst As New ADODB.Recordset
    'Dim cnn As ADODB.Connection
    'Dim strColorID As String
    Dim errMsg$

    errMsg = "Could not find Base." & vbCrLf & "Check design code extension"

    If CurrentProject.AllForms("frmCreateItem").IsLoaded = False Then
        Exit Sub
    End If

    'If no base is found
    If IsNull(Me.Base.ItemData(1)) And Me.Armor <> -1 Then
        'Check if construction is a loose tube design
        If Me.ItemNo Like "L*" Then
            If MsgBox("Is this a loose tube armor build?", vbYesNo) = vbYes Then
                If Not LooseTube_Armor(Me.ItemNo) Then
                    MsgBox (errMsg), vbCritical
                    Exit Sub
                End If
            End If
        Else
            MsgBox (errMsg), vbCritical
            Exit Sub
        End If
    End If



    'This sub routine completes the base, jacket, color, colorchip and fiber type for an item with one base

    If IsNull(Me.Base.ItemData(2)) = False And Me.Armor <> -1 Then


        MsgBox ("Multiple Possible Constructions.  Review"), vbCritical


    ElseIf Me.Armor <> -1 Then

   
        Me.Base.Requery
        Me.Base = Me.Base.ItemData(1)
    
    End If

    Call AfterBaseSelection
    
    


    
End Sub

Public Sub Base_AfterUpdate()
    
    On Error GoTo Err_Base_AfterUpdate
    Call AfterBaseSelection
   
    
Exit_Base_AfterUpdate:

    
    Exit Sub
    
    
Err_Base_AfterUpdate:

    
    MsgBox Err.Description
    Resume Exit_Base_AfterUpdate
    
End Sub

Sub BronzeBraided()


    bBronzeBraided = True


End Sub

Private Sub btnAttachSpec_Click()
    ' Requires reference to Microsoft Office 11.0 Object Library.
 
    Dim fDialog As Office.FileDialog
    Dim varFile As Variant
 
    ' Clear listbox contents.
    Me.FileList.rowSource = ""
 
    ' Set up the File Dialog.
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
 
    With fDialog
 
        ' Allow user to make multiple selections in dialog box
        .AllowMultiSelect = True
             
        ' Set the title of the dialog box.
        .Title = "Please select one or more files"
 
        ' Clear out the current filters, and add our own.
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .Filters.Add "Access Databases", "*.MDB"
        .Filters.Add "Access Projects", "*.ADP"
      
 
        ' Show the dialog box. If the .Show method returns True, the
        ' user picked at least one file. If the .Show method returns
        ' False, the user clicked Cancel.
        If .Show = True Then
 
            'Loop through each file selected and add it to our list box.
         
            For Each varFile In .SelectedItems
                Me.FileList.AddItem varFile
            Next
 
        Else
            'MsgBox "You clicked Cancel in the file dialog box."
        End If
    End With


End Sub

Public Sub btnClose_Click()
    On Error GoTo Err_cmdCloseForm_Click
    
    'DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close acForm, "frmCreateItem"
    Forms!frmMainMenu.Form.Visible = True
    
Exit_cmdCloseForm_Click:
    Exit Sub
    
Err_cmdCloseForm_Click:
    MsgBox Err.Description
    Resume Exit_cmdCloseForm_Click
    
    
End Sub

Private Sub btnFindSub_Click()
    Dim iComboItem As Integer

    For iComboItem = 1 To Me.cbSubunit.ListCount - 1
        Call OpenAClient(3, Me.cbSubunit.ItemData(iComboItem), 1)
    Next


End Sub

Private Sub btnFind_Click()
    Dim iComboItem As Integer

    For iComboItem = 1 To Me.Base.ListCount - 1
        CableConstruction.Load (Me.Base.ItemData(iComboItem))
    
        Call OpenAClient(3, Me.Base.ItemData(iComboItem), Nz(CableConstruction.Location, 2))
    Next


End Sub

Private Sub btnPrint_Click()
    DoCmd.OpenForm "frmPrintLineView"
End Sub

Private Sub btnSaveSpec_Click()
    Dim CustomerSpec As String
    Dim SpecRev As String
    Dim Customer As String
    Dim strCustomerPartNumber As String
    Dim SpecPath As String

    Me.CustomerRev = UCase(Me.CustomerRev)

    CustomerSpec = UCase(Me.Spec)
    SpecRev = Me.CustomerRev
    Customer = Me.Customer
    strCustomerPartNumber = Me.PID




    If Me.FileList.ListCount <> 0 Then
        SpecPath = Forms.frmCreateItem.FileList.ItemData(0)
        Call SaveCustomerSpec(CustomerSpec, SpecRev, Customer, SpecPath)
    End If

    Call SaveSpectoTable(Customer, CustomerSpec, strCustomerPartNumber)


End Sub

Private Sub btnShowSpec_Click()

    Dim Customer As String
    Dim CustomerSpec As String

    Customer = Me.Customer
    CustomerSpec = Me.Spec

    Call ShowCustomerSpec(CustomerSpec, Customer)


End Sub

Private Sub btnSubmit_Click()
    Dim msgresponse As Integer
    Dim stDocName As String
    Dim strSearch As String
    Dim ExistingItem As New clsCableQualities

    
    If ExistingItem.Load(Me.NewOracle_) Then
        MsgBox "This Oracle part number already exists. Please review.", vbCritical
        DoCmd.OpenForm "frmFindCustomerPID", acNormal, , "Oracle = '" & Me.NewOracle_ & "'", , acWindowNormal
        Me.btnSubmit.Enabled = False
        Exit Sub
    End If

    On Error GoTo Err_btnSubmit_Click
    
    If IsNull(Me.Color) Then
        MsgBox "Select a Color"
        Exit Sub
    End If
    
    If IsNull(getColorChip(Me.Color, Me.Jacket)) Then
        MsgBox "No color chip found for this color jacket combination" & vbCrLf & vbCrLf & "A color chip must be added before you can continue", vbCritical
        Exit Sub
    End If
        
    
    If IsNull(Me.Color) = True Or IsNull(Me.Base) = True Or IsNull(Me.Fiber1) = True _
                                                                                Or IsNull(Me.Customer) = True Then
        MsgBox "Fill out form completely", vbCritical
        Exit Sub
    End If
    
    
    If IsNull(Me.cbSubunit) And Me.Location = 2 And Me.Armor <> -1 And bBronzeBraided = False Then
        MsgBox "Subunit construction is missing", vbCritical
        Exit Sub
    End If
    
    If Me.FileList.ListCount <> 0 And (IsNull(Me.PID) = True Or IsNull(Me.Spec) = True Or IsNull(Me.CustomerRev) = True) Then
        MsgBox "Fill out SpecDoc#, PID, and/or Rev", vbCritical
        Exit Sub
    End If
    
    'Identifies if a ribbon item is being made
    'Does this ribbon item need to be high Cure?
    If CableConstruction.TBTypeID = 8 And Mid(Me.ItemNo, 12, 1) <> "8" Then
        'Removed question to prompt for high cure because all Non-peelable ribbons have moved to High Cure
        'If MsgBox("Does this fiber need High Cure? " & vbCrLf & "Will this ribbon go into a TE, Tyco, or Commscope item?", vbYesNo + vbCritical) = vbYes Then
        CableConstruction.RibbonHighCure = True
        'End Select
        'End If
    End If
    
    'Set the listing label to the appropriate value
    'This value will be loaded into Oracle and interpreted at time of label print
    If Not IsNull(Me.ListingCompany) Then
        CreateNewItem.CableQualities.setListingLabel Me.ListingCompany, Not IsNull(Me.ListingType), Not IsNull(Me.ListingTypeCanada)
    End If
    
    
    
    If Me.Armor = -1 Then
        stDocName = "mcrbtnCreateItemSubmit"
    End If
    
    
    If IsNull(Me.ListingCompany) = True And Not (Me.NewOracle_) Like "pt*" And Mid(Me.ItemNo, 9, 1) <> "U" Then
        
        MsgBox ("Please choose a listing company before continuing."), vbCritical
        Exit Sub
        
    End If
        
    If Me.PrintType = 2 And Me.FileList.ListCount = 0 And IsNull(Me.PID) = False Then
        If MsgBox("Does the customer have a spec sheet?  Attach and resubmit.", vbYesNo + vbCritical, "Customer Spec") = vbYes Then
            Call btnAttachSpec_Click
            Exit Sub
        End If
    End If
    
    If Me.MixedFiber = 1 And IsNull(Me.Fiber2) Then
        MsgBox ("Fiber 2 is empty.  Ensure each mixed fiber has a fiber type." & vbCrLf & "Example: -18MM6/6SM9"), vbCritical
        Exit Sub
    End If

    'Check to see if the item possibly exists
    'Need to find a good position to perform this check and then continue with build
    '    If CreateNewItem.CatalogExist(Me.Customer, Me.PID, Me.Base, Me.Fiber1, Me.Fiber2, Me.ItemNo, CableConstruction.RibbonHighCure) Then
    '        Exit Sub
    '    End If
    '
        
    DoCmd.SetWarnings False
    
    msgresponse = MsgBox("Are you ready to create this item?", vbYesNo, "Continue")
    If msgresponse = vbNo Then
        Exit Sub
    End If
    
    'If Me.Armor <> -1 Then
    
    'bTBType = Me.Base.Column(7)
    
    'End If
    

    
    DoCmd.SetWarnings False
    If Me.Armor = 0 Then
        Select Case CableConstruction.DesignTypeID
        Case 36, 27
        
            strSearch = "[Fiber] = '" & Me.Fiber1 & "' AND [Base] = '" & Me.Base & "' AND mid([Item No],6,1) = '" & Mid(ItemNo, 6, 1) & "'AND mid([Item No],12,1) = '" & Mid(ItemNo, 12, 1) & "' "
        
        Case Else
            strSearch = "[Fiber2] = '" & Me.Fiber2 & "' AND [Fiber] = '" & Me.Fiber1 & "' AND [Base] ='" & Me.Base & "' AND [Color] = '" & Me.Color & _
                        "' AND RibbonHighCure = " & CBool(Int(CableConstruction.RibbonHighCure)) & _
                        "  AND left([Item No],8) = '" & Left(Me.ItemNo, 8) & "'  AND mid([Item No],10,3) = '" & Mid(Me.ItemNo, 10, 3) & "' AND mid([Item No],9,1) = 'U'"
            
            
        End Select
        
        DoCmd.OpenForm "frmFiberMatchExisting", acNormal, , strSearch
    Else
        Call CreateItem
        
    End If
    
    

    'if No file is attached then don't launch the save spec routine
    If IsNull(Me.Spec) = False Then
        Call btnSaveSpec_Click
    End If

    DoCmd.SetWarnings True

    DoCmd.Close acForm, "frmWait"

Exit_btnSubmit_Click:
    Exit Sub
    
Err_btnSubmit_Click:
    MsgBox Err.Description
    DoCmd.SetWarnings False
    Resume Exit_btnSubmit_Click
    
    
    
    
End Sub

Private Sub btnUnitSeries_Click()
    On Error GoTo Err_btnUnitSeries_Click
    
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim lResponse As Integer
    
    
    
    If Forms!frmCreateItem!Exists = -1 Then
        lResponse = MsgBox("Item already exists.  No need for a new Unit Series.", vbOKOnly, "Unit Series")
        
        
        
        
    ElseIf Forms!frmCreateItem!Exists = 0 Then
        stDocName = "frmUnitSeries"
        DoCmd.OpenForm stDocName, , , stLinkCriteria
        Forms!frmUnitSeries!NewUnit = Me.UnitSeries
        
    End If
    
Exit_btnUnitSeries_Click:
    Exit Sub
    
Err_btnUnitSeries_Click:
    MsgBox Err.Description
    Resume Exit_btnUnitSeries_Click
    
    
    
    
End Sub

Private Sub CableAttributes()

    Dim SubColor As String

    SubColor = Nz(Mid(Me.ItemNo, 11, 1))

    If CurrentProject.AllForms("frmFiberMatchExisting").IsLoaded = False And Me.ItemNo Like "*AIA*" Then ' Must Create core cable before creating armored item
    
        MsgBox ("Create core cable before creating armored item.")
    
        DoCmd.Close acForm, "frmCreateItem"
        Forms!frmMainMenu.Visible = True
        Exit Sub
    
    ElseIf bCenterMember = True Then
        Exit Sub


    
    End If
    
    
    Me.FiberType = Me.Fiber1Type
  

    If SubColor = "0" And Left(Me.ItemNo, 2) <> "KP" Then
        Forms!frmCreateItem!SubColor = 6
    ElseIf SubColor = "0" And Left(Me.ItemNo, 2) = "KP" Then
        Forms!frmCreateItem!SubColor = Mid(Me.ItemNo, 10, 1)
    Else
        Forms!frmCreateItem!SubColor = Mid(Me.ItemNo, 11, 1)
    End If
    
    
    Select Case SubColor
    Case "X"
        Forms!frmCreateItem!MixedSub = 1
        Forms!frmCreateItem!SubX = 1
        Forms!frmCreateItem!TBColor = 0
        If (Me.ItemNo Like "b*" Or Me.ItemNo Like "x*" Or Me.ItemNo Like "r*") Or Not Me.ItemNo Like "*/*" Then 'Identifies if the cable is a breakout and needs sequentially colored subunits
            Forms!frmCreateItem!HybridSubColor = 0
        Else
            Forms!frmCreateItem!HybridSubColor = 1 'Identifies if the cable needs subunits colored by it's Fiber Mode
        End If
    Case "Z"
        Forms!frmCreateItem!MixedTB = 1
        Forms!frmCreateItem!TBColor = 1
        Forms!frmCreateItem!SubX = 1
        Forms!frmCreateItem!MixedSub = 1
            
        
    Case "Y"
        If Me.ItemNo Like "*/*" Then
            Forms!frmCreateItem!HybridSubColor = 1
        End If
        Forms!frmCreateItem!TBY = 1
        Forms!frmCreateItem!TBColor = 1
        Forms!frmCreateItem!SubX = 1
        Forms!frmCreateItem!SubColor = Mid(Me.ItemNo, 10, 1)
    Case "S"
        Forms!frmCreateItem!StripePrint = 1
        Forms!frmCreateItem!TBColor = "S"
        Forms!frmCreateItem!SubX = 1
        Forms!frmCreateItem!SubColor = Mid(Me.ItemNo, 10, 1)
        Forms!frmCreateItem!StripePrint = 1
        
    Case "U"
        Forms!frmCreateItem!TBColor = "U"
        Forms!frmCreateItem!SubX = 1
        Forms!frmCreateItem!MixedSub = 1
        Forms!frmCreateItem!SubColor = Mid(Me.ItemNo, 10, 1)
        Forms!frmCreateItem!TBY = 0
        
    Case Else
        Forms!frmCreateItem!MixedTB = 0
        Forms!frmCreateItem!TBColor = 0
    End Select
        

        

    'Sets tb color
    Call frmCreateItem_Hold
    
    If CableConstruction.TBTypeID = 5 And Me.Location = 1 Then
        bFurcation = True
    ElseIf (CableConstruction.TBTypeID = 32 Or CableConstruction.TBTypeID = 34) And Me.Location = -1 Then
        bCenterMember = True
    Else
        bFurcation = False
        bCenterMember = False
    End If
    
    Call BronzeBraided

End Sub

Public Sub Class1Fiber()

    'Finds the correct class 1 fiber and inserts it into the form
    DoCmd.OpenForm "frmFiberClass"
    Me.OM = Forms!frmFiberClass![OM#]
    Me.BIF = Forms!frmFiberClass!BIF
    Me.Fiber1 = Forms!frmFiberClass!Fiber
    Me.Fiber = Me.Fiber1
    DoCmd.Close acForm, "frmFiberClass"


End Sub

Private Sub Color_AfterUpdate()
    'Me.ColorChip = Me.Color.Column(2)


    'determines if white print is needed for the finished cable based on the cable color

    If Me.Color Like "*blue*" Or Me.Color Like "*black*" Or Mid(Me.ItemNo, 10, 1) = "8" Or Mid(Me.ItemNo, 10, 1) = "1" Then
        Me.PrintWhite = 1
    Else
        Me.PrintWhite = 0
    End If

End Sub

Private Sub Command192_Click()
    On Error GoTo Err_Command192_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command192_Click:
    Exit Sub

Err_Command192_Click:
    MsgBox Err.Description
    Resume Exit_Command192_Click
    
End Sub

Private Sub Command48_Click()
    On Error GoTo Err_Command48_Click
    
    Dim stDialStr As String
    Dim PrevCtl As control
    Const ERR_OBJNOTEXIST = 2467
    Const ERR_OBJNOTSET = 91
    Const ERR_CANTMOVE = 2483
    
    Set PrevCtl = Screen.PreviousControl
    
    If TypeOf PrevCtl Is TextBox Then
        stDialStr = IIf(VarType(PrevCtl) > V_NULL, PrevCtl, "")
    ElseIf TypeOf PrevCtl Is ListBox Then
        stDialStr = IIf(VarType(PrevCtl) > V_NULL, PrevCtl, "")
    ElseIf TypeOf PrevCtl Is ComboBox Then
        stDialStr = IIf(VarType(PrevCtl) > V_NULL, PrevCtl, "")
    Else
        stDialStr = ""
    End If
    
    Application.Run "utility.wlib_AutoDial", stDialStr
    
Exit_Command48_Click:
    Exit Sub
    
Err_Command48_Click:
    If (Err = ERR_OBJNOTEXIST) Or (Err = ERR_OBJNOTSET) Or (Err = ERR_CANTMOVE) Then
        Resume Next
    End If
    MsgBox Err.Description
    Resume Exit_Command48_Click
    
End Sub

Public Sub CreateItem()


    'On Error GoTo ErrorHandler

    DoCmd.OpenForm "frmWait", acNormal, "", "", , acNormal
    ' Clear Basic Cable Hold Table
    DoCmd.OpenQuery "qryNewBasicCableDelete", acViewNormal, acEdit
    ' Clear Basic Construction Hold Table
    DoCmd.OpenQuery "qryNewPrductConstructionHoldClear", acViewNormal, acEdit
    ' Clear Unit Series Hold Table
    DoCmd.OpenQuery "qrytblUnitSeriesCopyClear", acViewNormal, acEdit
    



    If Exists = -1 Or Forms!frmCreateItem!Armor = -1 Or Mid(Me.ItemNo, 9, 1) = "T" Then
        Me.Unprinted = 0
  
        Me.btnSubmit.Enabled = False
    
    End If

    If Me.Location = 1 Then
        If Me.NewOracle_ Like "pt*" Then
            Call CreateNewTB
        Else
            Call CreateSinglePassCable
        End If
            
    ElseIf Me.Location = 2 Or Forms!frmCreateItem!Armor = -1 Then
        Call mcrCreateBasicCableSheath
        
    End If
        
    Me.btnClose.SetFocus
    CreateNewItem.NewItemNumber = Me.NewOracle_

        
    

    DoCmd.Close acForm, "frmWait"
    
    'ErrorExit:
    'Exit Sub
    '
    'ErrorHandler:
    'MsgBox "An error occurred during item creation, please contact product engineering.", vbCritical
    'DoCmd.Close acForm, Me.Name
    'Forms!frmMainMenu.Visible = True
    'Resume ErrorExit




End Sub

Private Sub CreateNewTB()


    If Not CreateNewItem.CableQualities.IsPrinted Then
        Me.Unprinted = -1
    End If
    Call FindTB
    'If FindTB = False Then
    'Call CreateSinglePassCable
    'Exit Sub
    'End If
    
    DoCmd.OpenForm "frmFiberSpec"
    
    DoCmd.OpenQuery "qryNewPrductConstructionHoldClear"
    DoCmd.OpenQuery "qryNewPrductConstructionClear"
    DoCmd.OpenQuery "qryLOADNewTBs"
    DoCmd.OpenQuery "qryUpdateNewTBs"
    Call TBMaterial_Update
    DoCmd.OpenQuery "qryNewBasicConstructionAppend"
    
    DoCmd.Close acForm, "frmFiberSpec"
    DoCmd.OpenForm "frmNewSubunit"
    

End Sub

Public Function CreateSinglePassCable()
    If Left(Me.ItemNo, 2) = "FB" Or Left(Me.NewOracle_, 2) = "pt" Then
        Me.Unprinted = 0
    End If
    
    
    
    
    DoCmd.OpenForm "frmFiberSpec"
    DoCmd.OpenQuery "qryNewBasicConstructionClear"
    DoCmd.OpenQuery "qryNewBasicConstruction"
    DoCmd.OpenQuery "qryUpdateNewProductConstruciton"
    
    'Finds the appropriate tight buffer type depending on the cable family, jacket material, and tigh buffer indicator
    If CableConstruction.TBTypeID = 5 Then
        DoCmd.OpenQuery "qryfrmCreateItemFurcationUpdate" 'Furcation tube
    ElseIf CableConstruction.TightBuffered Then
        Call TBMaterial_Update                   'Tight buffers
    Else
        DoCmd.OpenQuery "qryfrmCreateItemUpdateTBMaterial_NoMaterial"
    End If

    
    If Me.PrintType <> 0 Or Left(Me.ItemNo, 2) = "FB" Or Left(Me.NewOracle_, 2) = "pt" Or Left(Me.NewOracle_, 3) = "rbn" Then
        DoCmd.OpenQuery "qryNewBasicConstructionAppend"
    End If
    'Finds if cable needs unprinted "-00" extensions. This does not apply to "PF", bare fibers, items
    If Me.Unprinted = -1 Then
        If Left(Me.NewOracle_, 3) = "rbn" Then
            DoCmd.RunSQL "UPDATE tblNewPrductConstruction SET [New Oracle Part #] = [New Oracle Part #] & '-OSP' , [Customer] = 'OUTSIDE PROCESSING'"
        Else
            DoCmd.OpenQuery "qryNewProductBasicUnprintedUpdate"
        End If
       
        DoCmd.OpenQuery "qryNewProductBasicUnprintedAppend"

    End If
    
    If Not CreateNewItem.CableQualities.IsPrinted And Forms!frmCreateItem!PrintType <> 0 Then
        DoCmd.OpenQuery "qryNewProductBasicUnprintedUpdateCustomer"
    
    End If
    
    '8/2/2016 ***Causing confusion to associates on floor with special instructions.  Special instructions removed to avoid issues****
    'If the ribbon requires High Cure, then update the table to reflect high cure, RibbonHighCure = 1
    If CableConstruction.RibbonHighCure = True Then
        DoCmd.RunSQL "UPDATE tblNewProductConstructionHold SET RibbonHighCure = -1 " ',[Special Instr Product1] = 'CUSTOMER REQUIRES **HIGH CURE** RIBBONIZATION'"
    End If
    
    
    DoCmd.Close acForm, "frmFiberSpec"
    DoCmd.OpenForm "frmNewSubunit"
End Function

Private Sub Customer_AfterUpdate()
    Dim lResponse As Integer
    'Reset associated fields
    
    Dim customerHold As New clsCustomer
    
    
    If IsNull(Me.Customer) Then
        MsgBox "A customer must be input", vbCritical
        Me.Customer.SetFocus
        Exit Sub
    End If
    
    Me.PID = Null
    Me.PID.Requery
    '    Me.PrintLine1 = Null
    Me.PrintLine1.Requery
    '    Me.PrintLine2 = Null
    Me.PrintLine2.Requery
    If Me.Armor = 0 Then
        Me.PrintLine3 = Null
        Me.PrintLine3.Requery
    End If
    Me.Customer = Replace(Me.Customer, Chr(13) & Chr(10), "")
    
    
    If Me.Customer Like "*BELDEN*" Then
        'Me.ListingNumber = "BELDEN-D"
    ElseIf Me.Customer Like "*MOHAWK*" Then
        'Me.ListingNumber = "MOHAWK-D"
    ElseIf Me.Customer Like "*DURA*" Then
        Me.PrintLine1 = "AFL eABF OPTICAL CABLE"
        Me.PrintLine4 = "MM/YY 000000 FEET     (Print every other foot)"
    End If
    
    If (Me.Customer Like "*te con*" Or Me.Customer Like "*tyco*" Or Me.Customer Like "*commscope*") And (Me.OM = "OM3" Or Me.OM = "OM4") Then
        lResponse = MsgBox("Is this a Class 1 Fiber?", vbYesNo + vbCritical)
        If lResponse = vbYes Then
            
            Call Class1Fiber
                
        End If
    End If
    
    
    customerHold.Load (Me.Customer)
    
    If customerHold.IsExistingCustomer = False Then
        If MsgBox("Customer does not exist.  Do you want to add this customer?", vbYesNo) = vbYes Then
            customerHold.AddNewCustomer (Me.Customer)
            Set ItemCustomer = customerHold
        Else
            Me.Customer = ItemCustomer.Customer
        End If
    Else
        Set ItemCustomer = customerHold
        If Not IsNull(Me.Base) Then
            If ItemCustomer.IsProhibitedDesign(Me.Base) Then
                MsgBox "This base is prohibited for " & Me.Customer & "." & vbCrLf & " Comments why: " & ItemCustomer.CustomerProhibitedNotes, vbCritical
                Me.btnSubmit.Enabled = False
            Else
                'Me.btnSubmit.Enabled = True
            End If
        End If
            
    End If
    
    'If ItemCustomer.getCustomerNotes <> Empty Then
    Me.txCustomerNotes = ItemCustomer.getCustomerNotes
    'End If
        

    
    
    
End Sub

Private Sub Customer_GotFocus()
    '    Me.Customer.SetFocus
End Sub

Public Sub Fiber1_AfterUpdate()


    Me.Fiber = Me.Fiber1
    If IsNull(Me.Fiber1) = False Then
        mlFiber1.Load (Fiber1)
        Me.OM = mlFiber1.OM
        Me.BIF = mlFiber1.BIF
        Me.Mode = mlFiber1.Mode
    End If
    'DoCmd.OpenForm "frmFiberMode"
    

    
    If Me.MixedFiber = 1 Then
        
    Else
        
        If Me.MixedFiber = 0 Then
            mlFiber1.setFiberCount = Int(Mid(Me.ItemNo, 3, 3))
            

        Else
            mlFiber1.setFiberCount = Me.Fiber1Count
            Me.PrintLine2 = Null
            Me.PrintLine2.Requery
        End If
    End If
    'DoCmd.Close acForm, "frmFiberMode"
    
    If Me.PrintType = 3 And Me.MixedFiber = 0 Then
        PrintType_AfterUpdate
    End If
    
    Call FiberPrint
    
    'DoCmd.OpenQuery "qryfrmCreateItemFiberModeUpdate"
    'DoCmd.Requery "OM"
    'DoCmd.Requery "BIF"
End Sub

Private Sub Fiber1_GotFocus()

    If Me.MixedFiber <> 1 Then


        Select Case Me.FiberExt
        Case 1
            Me.Fiber1.rowSource = "qryfrmCreateItemFiber1TypeExtension"
        Case 2
            Me.Fiber1.rowSource = "qryfrmCreateItemFiber1TypeBIF"
            '        Case 3
            '            Me.Fiber1.RowSource = "qryfrmCreateItemFiber1TypeITU"
        Case Else
            Me.Fiber1.rowSource = "qryfrmCreateItemFiberType"
        End Select
    Else

        Select Case Me.FiberExt
        Case 1
            Me.Fiber1.rowSource = "qryfrmCreateItemFiber1TypeExtension"
        Case 2
            Me.Fiber1.rowSource = "qryfrmCreateItemFiber1TypeBIF"
            Me.Fiber2.rowSource = "qryfrmCreateItemFiber2TypeBIF"
        Case Else
            Me.Fiber1.rowSource = "qryfrmCreateItemFiberType"
            Me.Fiber2.rowSource = "qryfrmCreateItemFiber2Type"
            
        End Select
        
    End If
    Call UserCheck

    If sUser = 1 Then
        Me.Fiber1.LimitToList = False
        Me.Fiber2.LimitToList = False
    Else
        Me.Fiber1.LimitToList = True
        Me.Fiber2.LimitToList = True
    End If

End Sub

Private Sub Fiber2_AfterUpdate()
    
    On Error GoTo Err_Fiber2_AfterUpdate
    
    If IsNull(Me.Fiber2) = False Then
        mlFiber2.Load (Me.Fiber2)
        Me.OM2 = mlFiber2.OM
        Me.BIF2 = mlFiber2.BIF
        Me.Mode2 = mlFiber2.Mode
    End If
    
    Call FiberPrint

    If Me.PrintType = 3 Then
        PrintType_AfterUpdate
    End If

    
    Me.FiberType2 = Me.Fiber2
   
Exit_Fiber2:
    Exit Sub
    
Err_Fiber2_AfterUpdate:
    If Err.Number = 13 Then
        MsgBox "Mixed fiber counts must be at least 2 digits " & vbCrLf & "Example: -01MM5/01SM4"
    Else
        MsgBox Err.Description & " " & Err.Number
    End If
    'DoCmd.Close acForm, "frmFiberMode2"
    'DoCmd.Close acForm, "frmFiberMode"
    Resume Exit_Fiber2

    
 
    
End Sub

Private Sub FiberPrint()
    Dim cFiber As New Collection
    Dim i, j As Integer
    Dim PrintLine2$
    j = 1


    cFiber.Add mlFiber1
    cFiber.Add mlFiber2

    For i = 1 To cFiber.Count
        If cFiber(j).Fiber = "FBR00100" Or cFiber(j).Fiber = Empty Then
            cFiber.Remove j
            j = j - 1
        End If
        j = j + 1
    Next

    If cFiber.Count = 0 Then
        Exit Sub
    End If

    PrintLine2 = getFiberPrint(cFiber, mlDesignCode, CableConstruction.TBTypeID)

    Me.Fiber1Count = mlFiber1.getFiberCount
    Me.Fiber2Count = mlFiber2.getFiberCount

    If Not CreateNewItem.CableQualities.IsPrinted Then
        Exit Sub
    End If

    Me.PrintLine2 = PrintLine2


    If mFiberPrint.isVendorSpecificPrint And Me.PrintType = 3 Then
        Me.PID = UCase(mFiberPrint.FiberVendor) & " PRINT"
    End If

    If (Me.Jacket = "PVC00145" Or Me.Jacket = "JKT00120" Or Me.Jacket = "JKT00119") And Me.Color = "BLACK" And (Left(Me.ItemNo, 1) = "K" Or Left(Me.ItemNo, 1) = "Q") Then
        Me.PrintLine2 = Me.PrintLine2 & " SUN RES"
    End If



End Sub

Private Sub FiberSpecs2_Click()
    DoCmd.OpenForm "frmFiberSpec2View"
End Sub

Private Sub FiberSpecs_Click()
    On Error GoTo Err_FiberSpecs_Click
    
    Dim stDocName As String
    stDocName = "frmFiberSpecView"
    DoCmd.OpenForm stDocName
    
Exit_FiberSpecs_Click:
    Exit Sub
    
Err_FiberSpecs_Click:
    MsgBox Err.Description
    Resume Exit_FiberSpecs_Click
    
End Sub

Private Sub Form_Load()
    DoCmd.Maximize
    Dim Count As Integer
    Count = 1
    mlDesignCode = Me.ItemNo
    bFurcation = False
    bCenterMember = False
    FillerTube = False
    
    frmMainMenu.Visible = False

    If mlDesignCode Like "M?*@*" Then
        bCenterMember = True
    ElseIf mlDesignCode Like "*000@*" Then
        bFurcation = True
    End If
    
    If mlDesignCode Like "??000@??[1U]?00*" Then
        If MsgBox("Is this a filler?", vbYesNo) = vbYes Then
            FillerTube = True
            bCenterMember = False
            bFurcation = False
        End If
    End If
  

    
    If Mid(mlDesignCode, 9, 1) <> "U" Then
        CreateNewItem.CableQualities.IsPrinted = True
    Else
        CreateNewItem.CableQualities.IsPrinted = False
    End If
    
    'Determine cable color indicator
    If Len(mlDesignCode) >= 10 Then
        CableColorIndicator = Mid(mlDesignCode, 10, 1)
    Else
        CableColorIndicator = "0"
    End If

    Me.btnSubmit.Enabled = True
    
    Me.Fiber12 = 1
    
    Me.PrintType.rowSource = "qryfrmFiberType"
    
    Me.FileList.rowSource = ""
    
    Me.Color.rowSource = "qryfrmCreateItemChipColor"
    


    If sUser <> 1 Then
        Me.Fiber1.LimitToList = True
        Me.Fiber2.LimitToList = True
    End If
    
    Call SetBaseRowSource
    
    Call LockFormFields
    
End Sub

Sub SetBaseRowSource()
    Dim rowSource$
    Dim rowSourceCriteria$

    If CreateNewItem.ItemExtensionDesignCheck(mlDesignCode) Then
        rowSourceCriteria = " WHERE DesignExtension = '" & CreateNewItem.GetDesignCodeExtension & "'"
    Else
        rowSourceCriteria = " WHERE DesignExtension is null"
    End If

    If Mid(mlDesignCode, 9, 1) = "T" Then
        Me.Unprinted = 0
        rowSource = "qryfrmCreateItemBaseTestItems"
    ElseIf bFurcation Then
        rowSource = "qryfrmCreateItemBaseFurcation"
    ElseIf bCenterMember Then
        rowSource = "qryfrmCreateItemBaseCenterMember"
    Else
        rowSource = "qryfrmCreateItemBaseNew"
    End If
    
    rowSource = "SELECT * FROM " & rowSource & rowSourceCriteria
    
    
    Me.Base.rowSource = rowSource
    
End Sub

Sub LockFormFields()
    
    Forms!frmCreateItem!Base.Enabled = True
    Forms!frmCreateItem!Jacket.Enabled = True
    Forms!frmCreateItem!Color.Enabled = True
    Forms!frmCreateItem!ColorChip.Enabled = True
    Forms!frmCreateItem!Fiber1.Enabled = True

    
    If Me.ItemNo Like "*/*" Then
        Forms!frmCreateItem!MixedFiber = 1
        Forms!frmCreateItem!Fiber2.Visible = True
        Forms!frmCreateItem!OM2.Visible = True
        Forms!frmCreateItem!BIF2.Visible = True
        Forms!frmCreateItem!SubX = 1
        Forms!frmCreateItem!FiberSpecs2.Visible = True
        
    Else
        Forms!frmCreateItem!MixedFiber = 0
        Forms!frmCreateItem!Fiber2.Visible = False
        Forms!frmCreateItem!OM2.Visible = False
        Forms!frmCreateItem!BIF2.Visible = False
        Forms!frmCreateItem!FiberSpecs2.Visible = False
    End If
    
    
    
    
    
    Me.ListingCompany.Enabled = True
    Me.ListingCompany.Locked = False
    
    
    'Set all fields to enabled.  The main menu form will turn off certain fields depending on print
    Forms!frmCreateItem!Customer.Enabled = True
    Forms!frmCreateItem!PID.Enabled = True
    Forms!frmCreateItem!Spec.Enabled = True
    Forms!frmCreateItem!CustomerRev.Enabled = True
    Forms!frmCreateItem!ListingCompany.Enabled = True
    Forms!frmCreateItem!PrintLine1.Enabled = True
    Forms!frmCreateItem!PrintLine2.Enabled = True
    Forms!frmCreateItem!PrintLine3.Enabled = True
    Forms!frmCreateItem!PrintLine4.Enabled = True
    Forms!frmCreateItem!ListingType.Enabled = True
    Forms!frmCreateItem!ListingTypeCanada.Enabled = True
    
    


End Sub

Private Sub frmCreateItem_Hold()
    Dim TBType As String

    TBType = Mid(Me.ItemNo, 12, 1)

    Forms!frmCreateItem!TBType = TBType

End Sub

Private Sub ItemBuildChecks()

    If Mid(mlDesignCode, 10, 1) = "0" And (Me.Jacket <> "NONE" Or Me.Jacket = "") Then
        MsgBox ("Jacket color indicator cannot be '0'." & vbCrLf & "Insert a color designation from WI-123 document for the desired color."), vbCritical
        Me.btnSubmit.Enabled = False
    End If
    
    


End Sub

Private Sub Jacket_AfterUpdate()
    
    Me.Color = Null
    Me.ColorChip = Null
    Me.ColorChip.Requery
    Me.Color.Requery
    
    DoCmd.RunMacro "mcrfromCreateItemJacket"
    

    
    
    
End Sub

Private Sub KQ_TBMaterial()
    'KQ & KP series has different TB material requirements for differing fiber counts
    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Set cnn = CurrentProject.connection
    Dim rs As New ADODB.Recordset
    Dim strFiberCount As String

    rst.Open "tblNewPrductConstruction", cnn, adOpenKeyset, adLockOptimistic
    strFiberCount = Mid(rst![Item No], 3, 3)


    rs.Open "tblKQ_TBType WHERE FiberCount = '" & strFiberCount & "'", cnn, adOpenKeyset, adLockOptimistic

    rst![TB Material] = rs!TBMaterial
    rst.Update

    rst.Close
    rs.Close
    cnn.Close


End Sub

Public Sub ListingAutoFill()
    Me.ListingType = Me.ListingType.ItemData(0)
    Me.ListingConstruction = Me.ListingType.Column(1)
    Me.Listing = Me.ListingType.Column(3)
    Me.ListingSection = Me.ListingType.Column(2)



End Sub

Private Sub ListingCompany_AfterUpdate()
    DoCmd.RunMacro "mcroListingCompany"
    If Me.ListingCompany = "NONE" Then
        Me.ListingType.Enabled = False
        Me.ListingTypeCanada.Enabled = False
        Exit Sub
    Else
        Me.ListingType.Enabled = True
        Me.ListingTypeCanada.Enabled = False
    End If
    
    
    '    If Me.ListingCompany = "ETL" Then
    '        Me.ListingNumber = "115304"
    '    ElseIf Me.ListingCompany = "UL" Then
    '        Me.ListingNumber = "E121250"
    '    Else
    '        Me.ListingNumber = Null
        
    '    End If
    
    If IsNull(Me.PrintLine3) = True Or Me.PrintLine3.Enabled = True Then
        Me.PrintLine3 = Null
        Me.PrintLine3.Requery
    Else
    End If
    
    Me.ListingType = Null
    Me.ListingType.Requery
    Me.ListingTypeCanada = Null
    Me.ListingTypeCanada.Requery
    
    
    
End Sub

Private Sub ListingCompany_GotFocus()
    'If Mid(Me.ItemNo, 6, 1) = "@" Then
    'Me.ListingCompany.RowSourceType = "Value List"
    'Me.ListingCompany.RowSource = "NONE"
    'End If
End Sub

Private Sub ListingTypeCanada_AfterUpdate()

    If IsNull(Me.ListingTypeCanada) Then
        Exit Sub
    End If

    If Me.Customer Like "*belden*" Then
        Me.PrintLine3 = "c(" & Me.ListingCompany & ") TYPE " & Me.ListingType & " " & Me.ListingTypeCanada & " (" & Me.ListingCompany & ") " & Me.ListingType & " RoHS"

    ElseIf Me.Customer Like "TE CON*" Or Me.Customer Like "TYCO*" Or Me.Customer Like "*COMMSCOPE*" Then
        Me.PrintLine3 = Me.ListingType & "-" & Me.ListingTypeCanada & " c(" & Me.ListingCompany & ")us " & Me.ListingType & " RoHS"

    Else
        Me.PrintLine3 = "(" & Me.ListingCompany & ") TYPE " & Me.ListingType & " " & Me.ListingTypeCanada & " c(" & Me.ListingCompany & ") RoHS"

    End If

    Call PrintLine3_AfterUpdate

End Sub

Public Sub ListingType_AfterUpdate()

    Me.ListingConstruction = Me.ListingType.Column(1)
    Me.Listing = Me.ListingType.Column(3)
    Me.ListingSection = Me.ListingType.Column(2)

    If Me.ListingCompany = "MSHA" And Me.ItemNo Like "*aia*" Then

        If Not Me.PrintLine2 Like "*" & Me.ListingType & "*" Then
    
            Me.PrintLine2 = Me.PrintLine2 & "    " & Me.ListingType
            Me.PrintLine3 = "CAUTION! Not to be used where permissible equipment is required"
        End If
    
        Exit Sub
    ElseIf Me.ListingCompany = "MSHA" Then
        If Not Me.PrintLine3 Like "*" & Me.ListingType & "*" Or IsNull(Me.PrintLine3) = True Then
            Me.PrintLine3 = Me.ListingType
        End If
    
        Exit Sub
    End If

    If Me.Customer Like "*belden*" Then
        Me.PrintLine3 = "c(" & Me.ListingCompany & ") TYPE " & Me.ListingType & " (" & Me.ListingCompany & ") " & Me.ListingType & " RoHS"

    ElseIf Me.Customer Like "TE C*" Or Me.Customer Like "TYCO*" Or Me.Customer Like "*COMMSCOPE*" Then
        Me.PrintLine3 = Me.ListingType & " c(" & Me.ListingCompany & ")us " & Me.ListingType & " RoHS"

    Else
        Me.PrintLine3 = "(" & Me.ListingCompany & ") TYPE " & Me.ListingType & "  RoHS"
    
    End If

    Call PrintLine3_AfterUpdate
    Me.ListingTypeCanada = Null
    Me.ListingTypeCanada.Requery
    Me.ListingTypeCanada = Me.ListingTypeCanada.ItemData(0)

    Call ListingTypeCanada_AfterUpdate


End Sub

Sub ListingUpdate()
    Dim ENumber As String, printLine As String

    printLine = Nz(Me.PrintLine1, "") + " " + Nz(Me.PrintLine2, "") + " " + Nz(Me.PrintLine3, "")

    'If Customer ID is empty then requery and see if the customer ID can be gathered
    If (ItemCustomer.CustomerID = "") Then ItemCustomer.Load (Me.Customer)

    'If the customer ID is still empty then run an update
    If (ItemCustomer.CustomerID = "") Then
    
        Me.Customer.Requery
        Customer_AfterUpdate

    End If

    ENumber = BurnListing.getBurnFileNumber(ItemCustomer.CustomerID, Me.ListingType.Column(4))

    Me.ListingNumber = ENumber

    'If the customer is AFL or the print contains AFL then the eNumber is not needed within the print
    If printLine Like "*afl *" Then
        
        ENumber = ""
    
    End If
    
    
    If ENumber <> "" And Not printLine Like "*" & ENumber & "*" Then
        Me.PrintLine3 = ENumber & " " & Me.PrintLine3
    End If
    
    
End Sub

Private Sub m_frmMainMenu_ItemRequestNumber(value As String)
    m_ItemRequestNumber = value
    CreateNewItem.ItemRequestNumber = value
End Sub

Private Sub PID_AfterUpdate()

    If IsNull(Me.PID) Then
        Exit Sub
    End If
    Me.PID = Replace(Me.PID, Chr(13) & Chr(10), "")

    If Me.PrintType = 2 Then
        If CommonFunctions.DoesPIDExist(Me.Customer, Me.PID) = True Then
            MsgBox ("There is an active Item for this PID and Customer." & vbCrLf & vbCrLf & "Either deactivate the existing item" & vbCrLf & "Or" & vbCrLf & "Update existing item Spec and Rev"), vbCritical
            Me.btnSubmit.Enabled = False
            DoCmd.OpenForm "frmFindCustomerPID", acNormal, , "Customer = '" & Me.Customer & "' AND [PID] = '" & Me.PID & "'", , acWindowNormal
        Else
            Me.btnSubmit.Enabled = True
        End If
    
    End If


End Sub

Private Sub PrintLine3_AfterUpdate()
    If IsNull(Me.Customer) Then
        MsgBox "Must choose a Customer", vbCritical
        Exit Sub
    End If

    Call ListingUpdate

    
    
End Sub

Private Sub PrintLine4_AfterUpdate()
    If Me.ItemNo Like "*test*" Or Mid(Me.ItemNo, 9, 1) = "2" Then
        Exit Sub
    End If

    If Me.PrintLine4.Column(1) = 1 Then
        If Me.ItemNo Like "*MSHA*" Then
            Me.NewOracle_ = Left(Me.NewOracle_, 8) & "03"
        Else
            Me.NewOracle_ = Left(Me.NewOracle_, 8) & "01"
        End If
    
    ElseIf Me.PrintLine4.Column(1) = 2 Then
        If Me.ItemNo Like "*MSHA*" Then
            Me.NewOracle_ = Left(Me.NewOracle_, 8) & "04"
        Else
            Me.NewOracle_ = Left(Me.NewOracle_, 8) & "02"
        End If
    
    End If
End Sub

Public Sub PrintType_AfterUpdate()
    DoCmd.RunMacro "mcrPrintType"
    Dim eAbfPrint, AflPrint, AflGenericPrint As String
    
    eAbfPrint = "AFL eABF OPTICAL CABLE"
    AflPrint = "AFL OPTICAL CABLE   1-800-AFL-FIBER"
    AflGenericPrint = "OPTICAL FIBER CABLE"
    
    If CreateNewItem.CableQualities.IsPrinted Then
        Select Case Me.PrintType
        Case 1
           
            If Left(Me.ItemNo, 1) = "A" Then
                Me.PrintLine1 = eAbfPrint
            Else
                Me.PrintLine1 = AflPrint
            End If
        Case 2
            If Left(Me.ItemNo, 1) = "A" Then
                Me.PrintLine1 = eAbfPrint
            Else
                Me.PrintLine1 = AflGenericPrint  'Creates Print Line 1
            End If
        Case 3
            Me.PrintLine1 = AflGenericPrint
        End Select
    End If
    


    'Dissables generic print PID, Doc, and Rev input
    If Me.PrintType = 2 Then
        Me.Customer.Enabled = True
        Me.PID.Enabled = True
        Me.Spec.Enabled = True
        Me.CustomerRev.Enabled = True
        Me.PID = Null
        'Me.PID.SetFocus
    Else
        Me.Customer.Enabled = False
        Me.PID.Enabled = False
        Me.Spec.Enabled = False
        Me.CustomerRev.Enabled = False
    End If

    If Me.PrintType = 3 Then
        If Me.ListingCompany.Enabled = True Then
            Me.ListingCompany.SetFocus
        End If
    Else
        If Me.Customer.Enabled = True Then
            Me.Customer.SetFocus
        End If
    End If

    If Me.PrintType = 2 And Not CreateNewItem.CableQualities.IsPrinted Then
        Me.NewOracle_ = Left(Me.NewOracle_, 7) & "-10"
    ElseIf Me.PrintType = 0 And Not CreateNewItem.CableQualities.IsPrinted Then
        Me.NewOracle_ = Left(Me.NewOracle_, 7) & "-00"
    End If

    If mFiberPrint.isVendorSpecificPrint And Me.PrintType = 3 Then
        Me.PID = mFiberPrint.FiberVendor & " PRINT"
    End If


        
End Sub

Private Sub PrintType_GotFocus()
    Dim PrintType As String

    PrintType = Mid(Me.ItemNo, 9, 1)
    Select Case PrintType
    Case 2
        Me.PrintType.rowSource = "qryfrmFiberTypeSpecialPrint"
    Case 1
        Me.PrintType.rowSource = "qryfrmFiberTypeStandard"
    Case "U"
        'me.PrintType.RowSource = ""
    End Select

End Sub

Private Sub Refresh_Click()
        
    Me.Requery
    
    
    DoCmd.OpenForm "frmBaseNomOD"
    Forms!frmCreateItem!NomOD = Forms!frmBaseNomOD!OD
    DoCmd.Close acForm, "frmBaseNomOD"
    
    
    
    
End Sub

Private Sub Spec_AfterUpdate()
    If IsNull(Me.Spec) = False Then
        Me.Spec = Replace(Me.Spec, Chr(13) & Chr(10), "")
    End If
End Sub

Public Sub StandardPrint()


    Call FiberPrint



End Sub

Private Sub StripedPrint_AfterUpdate()
    If Me.StripedPrint = -1 Then
        Me.StripePrint = 1
    Else
        Me.StripePrint = 0
    End If
End Sub

Public Sub TBMaterial_Update()




    If Left(Me.ItemNo, 2) = "WB" Or Left(Me.ItemNo, 2) = "RB" Then
        Exit Sub
    End If


    DoCmd.OpenForm "frmTBMaterial"





    If Me.NewOracle_ Like "pt*" Or Me.NewOracle_ Like "pf*" Then
        Forms!frmTBMaterial.RecordSource = "qryfrmCreateItemTBMaterialMatch_ForTB"
    Else
        Forms!frmTBMaterial.RecordSource = "qryfrmCreateItemTBMaterialMatch"
    End If
    'Select Case Me.TBType
    'Case 1, 2, 3, 4, 5, 6, 7, 8, 9, "G", "H", "U", "E"
    Forms!frmTBMaterial.Requery

    'If no TB material is located for the design then it notifies the user
    If IsNull(Forms!frmTBMaterial![TB Material]) = True Then
        MsgBox "No TB Material found for TB Type, Stop Item Creation", vbCritical
        'Exit Sub
    
    End If


    DoCmd.OpenQuery "qryfrmCreateItemTBMaterialUpdate"

    DoCmd.Close acForm, "frmTBMaterial"

    If Left(Me.ItemNo, 2) = "KQ" Or Left(Me.ItemNo, 2) = "KP" Then
        Call KQ_TBMaterial
    End If


    'End Select

End Sub

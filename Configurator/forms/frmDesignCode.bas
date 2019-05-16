Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    RecordLocks =2
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =2
    GridX =24
    GridY =24
    Width =20520
    DatasheetFontHeight =10
    ItemSuffix =106
    Left =5235
    Top =4665
    Right =27540
    Bottom =13650
    DatasheetGridlinesColor =12632256
    Filter ="[DesignTypeGroupID] like \"*\" AND SalesForceDescription like \"Simplex*\" AND F"
        "iberCount like \"1*\" AND [JacketType] like \"Furcation*\" AND [Base]  like \"SF"
        "000-3.0*\""
    OrderBy ="[vDesignCode].[SubOD]"
    RecSrcDt = Begin
        0x2505a02592d7e440
    End
    RecordSource ="vDesignCode"
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
        Begin OptionGroup
            SpecialEffect =3
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
            Height =5640
            BackColor =-2147483633
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3360
                    Top =1620
                    Width =720
                    Height =300
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="FiberCount"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeFiberCount"
                    ColumnWidths ="2520"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =1620
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1920
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1680
                            Top =1620
                            Width =1065
                            Height =300
                            Name ="Label17"
                            Caption ="Fiber Count"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =1620
                            LayoutCachedWidth =2745
                            LayoutCachedHeight =1920
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =3360
                    Top =360
                    Width =3660
                    Height =300
                    ColumnOrder =5
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Cable"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblDesignTypeGroups.DesignTypeGroupID, tblDesignTypeGroups.Desig"
                        "nTypeGroupDesc, tblDesignTypeGroups.Active FROM tblDesignTypeGroups WHERE (((tbl"
                        "DesignTypeGroups.Active)<>0)); "
                    ColumnWidths ="0;4320;360;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =3360
                    LayoutCachedTop =360
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1680
                            Top =360
                            Width =1080
                            Height =300
                            Name ="Label7"
                            Caption ="Cable Family"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =360
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3360
                    Top =780
                    Width =3660
                    Height =300
                    ColumnOrder =6
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"140\""
                    Name ="Jacket"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeJacketType"
                    ColumnWidths ="2520;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =3360
                    LayoutCachedTop =780
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1680
                            Top =780
                            Width =1065
                            Height =300
                            Name ="Label9"
                            Caption ="Cable Type"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =780
                            LayoutCachedWidth =2745
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =8
                    ListWidth =10080
                    Left =3360
                    Top =2040
                    Width =3660
                    Height =300
                    ColumnOrder =7
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";"
                        "\"10\";\"0\""
                    Name ="OD"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeBaseNomOD"
                    ColumnWidths ="0;0;720;1080;1296;3600;2160;720"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3360
                    LayoutCachedTop =2040
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =2340
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1680
                            Top =2040
                            Width =1065
                            Height =300
                            Name ="Label15"
                            Caption ="OD"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =2040
                            LayoutCachedWidth =2745
                            LayoutCachedHeight =2340
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3360
                    Top =3180
                    Width =1260
                    ColumnOrder =8
                    TabIndex =16
                    Name ="Base"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =3180
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =3420
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1680
                            Top =3180
                            Width =645
                            Height =240
                            Name ="txtSubOD"
                            Caption ="Base"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =3180
                            LayoutCachedWidth =2325
                            LayoutCachedHeight =3420
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15420
                    Top =1760
                    Width =2520
                    ColumnOrder =9
                    TabIndex =24
                    Name ="BaseOD"

                    LayoutCachedLeft =15420
                    LayoutCachedTop =1760
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =2000
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14280
                            Top =1760
                            Width =645
                            Height =240
                            Name ="txtOD"
                            Caption ="OD:"
                            LayoutCachedLeft =14280
                            LayoutCachedTop =1760
                            LayoutCachedWidth =14925
                            LayoutCachedHeight =2000
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3360
                    Top =2820
                    Width =3660
                    ColumnOrder =10
                    TabIndex =14
                    Name ="JacketMaterial"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =3060
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1680
                            Top =2820
                            Width =1260
                            Height =240
                            Name ="Label24"
                            Caption ="Jacket Material "
                            LayoutCachedLeft =1680
                            LayoutCachedTop =2820
                            LayoutCachedWidth =2940
                            LayoutCachedHeight =3060
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6180
                    Top =3180
                    Width =840
                    ColumnOrder =11
                    TabIndex =17
                    Name ="Location"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =3180
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =3420
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4980
                            Top =3180
                            Width =1020
                            Height =240
                            Name ="Label26"
                            Caption ="Cable Passes"
                            LayoutCachedLeft =4980
                            LayoutCachedTop =3180
                            LayoutCachedWidth =6000
                            LayoutCachedHeight =3420
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3360
                    Top =2460
                    Width =3660
                    ColumnOrder =12
                    TabIndex =12
                    Name ="Description"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =2460
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =2700
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1680
                            Top =2460
                            Width =930
                            Height =240
                            Name ="Label32"
                            Caption ="Description"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =2460
                            LayoutCachedWidth =2610
                            LayoutCachedHeight =2700
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9120
                    Top =360
                    Width =720
                    Height =300
                    ColumnOrder =13
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="Fiber1"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeFiber1Count2Pass"
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =360
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7440
                            Top =360
                            Width =1065
                            Height =300
                            Name ="Label3"
                            Caption ="Fiber1"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =360
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =9120
                    Top =2100
                    Width =3660
                    Height =300
                    ColumnOrder =16
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="TBColor"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeTBColor"
                    ColumnWidths ="0;3600"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =2100
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =2400
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7440
                            Top =2100
                            Width =1065
                            Height =300
                            Name ="txtTBColor"
                            Caption ="Sub Color"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =2100
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =2400
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =9120
                    Top =2520
                    Width =3660
                    Height =300
                    ColumnOrder =17
                    FontSize =10
                    FontWeight =700
                    TabIndex =15
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="PrintType"
                    RowSourceType ="Table/Query"
                    RowSource ="tblPrintType"
                    ColumnWidths ="0;3600"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =2520
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =2820
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7440
                            Top =2520
                            Width =1065
                            Height =300
                            Name ="txtPrintType"
                            Caption ="Print Type"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =2520
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =2820
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =12
                    ListWidth =4320
                    Left =9120
                    Top =780
                    Width =3660
                    Height =300
                    ColumnOrder =18
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BoundColumn =-1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="Fiber1Type"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeFiberType"
                    ColumnWidths ="0;1800;0;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =9120
                    LayoutCachedTop =780
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7440
                            Top =780
                            Width =1065
                            Height =300
                            Name ="Label38"
                            Caption ="Fiber1 Type"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =780
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =12
                    ListWidth =4320
                    Left =9120
                    Top =1620
                    Width =3660
                    Height =300
                    ColumnOrder =19
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    BoundColumn =-1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="Fiber2Type"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeFiberType"
                    ColumnWidths ="0;1800;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =9120
                    LayoutCachedTop =1620
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =1920
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =7440
                            Top =1620
                            Width =1065
                            Height =300
                            Name ="txtFiber2Type"
                            Caption ="Fiber2 Type"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =1620
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =1920
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15420
                    Top =1440
                    Width =2520
                    ColumnOrder =20
                    TabIndex =21
                    Name ="TB"

                    LayoutCachedLeft =15420
                    LayoutCachedTop =1440
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =1680
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14280
                            Top =1440
                            Width =660
                            Height =240
                            Name ="Label42"
                            Caption ="TB Type"
                            LayoutCachedLeft =14280
                            LayoutCachedTop =1440
                            LayoutCachedWidth =14940
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin Rectangle
                    BorderWidth =3
                    OverlapFlags =255
                    Left =1500
                    Top =180
                    Width =5700
                    Height =3660
                    Name ="Box43"
                    LayoutCachedLeft =1500
                    LayoutCachedTop =180
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =3840
                End
                Begin Rectangle
                    BorderWidth =3
                    OverlapFlags =255
                    Left =7320
                    Top =180
                    Width =5640
                    Height =3660
                    Name ="Box44"
                    LayoutCachedLeft =7320
                    LayoutCachedTop =180
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =3840
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =11820
                    Top =4200
                    ColumnOrder =21
                    TabIndex =19
                    Name ="Armor"

                    LayoutCachedLeft =11820
                    LayoutCachedTop =4200
                    LayoutCachedWidth =12080
                    LayoutCachedHeight =4440
                    Begin
                        Begin Label
                            OverlapFlags =127
                            Left =12060
                            Top =4200
                            Width =480
                            Height =240
                            Name ="Label46"
                            Caption ="Armor"
                            LayoutCachedLeft =12060
                            LayoutCachedTop =4200
                            LayoutCachedWidth =12540
                            LayoutCachedHeight =4440
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =14640
                    Top =2820
                    ColumnOrder =22
                    TabIndex =25
                    Name ="Corning"

                    LayoutCachedLeft =14640
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14900
                    LayoutCachedHeight =3060
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15420
                            Top =2820
                            Width =690
                            Height =240
                            Name ="Label61"
                            Caption ="Corning"
                            LayoutCachedLeft =15420
                            LayoutCachedTop =2820
                            LayoutCachedWidth =16110
                            LayoutCachedHeight =3060
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =11460
                    Top =4080
                    Width =1500
                    Height =480
                    Name ="Box62"
                    LayoutCachedLeft =11460
                    LayoutCachedTop =4080
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =4560
                End
                Begin ComboBox
                    OverlapFlags =255
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =9120
                    Top =1260
                    Width =3660
                    Height =300
                    ColumnOrder =23
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="TBType"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeTightBufferType"
                    ColumnWidths ="0;1800;0"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =9120
                    LayoutCachedTop =1260
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =7440
                            Top =1260
                            Width =1065
                            Height =300
                            Name ="txtTBType"
                            Caption ="TB Type"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =1260
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =1560
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15420
                    Top =800
                    Width =2520
                    ColumnOrder =24
                    TabIndex =18
                    Name ="ItemNoPrefix"

                    LayoutCachedLeft =15420
                    LayoutCachedTop =800
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =1040
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13740
                            Top =800
                            Width =1200
                            Height =240
                            Name ="Label66"
                            Caption ="ItemNoPrefix"
                            LayoutCachedLeft =13740
                            LayoutCachedTop =800
                            LayoutCachedWidth =14940
                            LayoutCachedHeight =1040
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11700
                    Top =4860
                    Width =1320
                    TabIndex =22
                    Name ="btnSubmit"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11700
                    LayoutCachedTop =4860
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =5220
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15420
                    Top =1120
                    Width =2520
                    ColumnWidth =3570
                    ColumnOrder =25
                    TabIndex =20
                    Name ="ItemNo"

                    LayoutCachedLeft =15420
                    LayoutCachedTop =1120
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =1360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13740
                            Top =1120
                            Width =1200
                            Height =240
                            Name ="Label69"
                            Caption ="Design Code"
                            LayoutCachedLeft =13740
                            LayoutCachedTop =1120
                            LayoutCachedWidth =14940
                            LayoutCachedHeight =1360
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13080
                    Top =840
                    Width =480
                    ColumnOrder =3
                    TabIndex =4
                    Name ="Mode1"

                    LayoutCachedLeft =13080
                    LayoutCachedTop =840
                    LayoutCachedWidth =13560
                    LayoutCachedHeight =1080
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13080
                    Top =1680
                    Width =480
                    ColumnOrder =26
                    TabIndex =11
                    Name ="Mode2"

                    LayoutCachedLeft =13080
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13560
                    LayoutCachedHeight =1920
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9120
                    Top =1200
                    Width =720
                    Height =300
                    ColumnOrder =14
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="Fiber2"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeFiber2Count2Pass"
                    ColumnWidths ="0;1800;1440"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =1200
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =7440
                            Top =1200
                            Width =1065
                            Height =300
                            Name ="txtFiber2"
                            Caption ="Fiber2"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =1200
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =1500
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15420
                    Top =2080
                    Width =2520
                    ColumnOrder =27
                    TabIndex =23
                    Name ="CableType"

                    LayoutCachedLeft =15420
                    LayoutCachedTop =2080
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =2320
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14100
                            Top =2080
                            Width =900
                            Height =240
                            Name ="Label80"
                            Caption ="CableType:"
                            LayoutCachedLeft =14100
                            LayoutCachedTop =2080
                            LayoutCachedWidth =15000
                            LayoutCachedHeight =2320
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =14160
                    Top =5340
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Base"
                    LayoutCachedLeft =14160
                    LayoutCachedTop =5340
                    LayoutCachedWidth =15780
                    LayoutCachedHeight =5640
                End
                Begin Label
                    OverlapFlags =85
                    Left =11280
                    Top =5340
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label85"
                    Caption ="Jacket Material"
                    LayoutCachedLeft =11280
                    LayoutCachedTop =5340
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =5640
                End
                Begin Label
                    OverlapFlags =85
                    Left =1320
                    Top =5340
                    Width =900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label86"
                    Caption ="Sub Fiber Count"
                    LayoutCachedLeft =1320
                    LayoutCachedTop =5340
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =5640
                End
                Begin Label
                    OverlapFlags =85
                    Left =4860
                    Top =5340
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Fiber Count"
                    LayoutCachedLeft =4860
                    LayoutCachedTop =5340
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =5640
                End
                Begin Label
                    OverlapFlags =85
                    Left =12960
                    Top =5340
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label33"
                    Caption ="Nominal OD"
                    LayoutCachedLeft =12960
                    LayoutCachedTop =5340
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =5640
                End
                Begin Label
                    OverlapFlags =85
                    Left =3660
                    Top =5340
                    Width =1140
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label39"
                    Caption ="# Subs Pos"
                    LayoutCachedLeft =3660
                    LayoutCachedTop =5340
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =5640
                End
                Begin Label
                    OverlapFlags =85
                    Left =6480
                    Top =5340
                    Width =4620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="Description"
                    LayoutCachedLeft =6480
                    LayoutCachedTop =5340
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =5640
                End
                Begin Label
                    OverlapFlags =85
                    Left =17040
                    Top =5340
                    Width =1200
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label43"
                    Caption ="TB Type"
                    LayoutCachedLeft =17040
                    LayoutCachedTop =5340
                    LayoutCachedWidth =18240
                    LayoutCachedHeight =5640
                End
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =5340
                    Width =960
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="Sub OD"
                    LayoutCachedLeft =240
                    LayoutCachedTop =5340
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =5640
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11100
                    Top =4860
                    Width =480
                    TabIndex =26
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

                    LayoutCachedLeft =11100
                    LayoutCachedTop =4860
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =5220
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =4800
                    Width =1320
                    TabIndex =27
                    Name ="btnSelect"
                    Caption ="Select"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =4800
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =5160
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15420
                    Top =480
                    Width =2520
                    ColumnOrder =2
                    TabIndex =28
                    Name ="SubFiberCount"

                    LayoutCachedLeft =15420
                    LayoutCachedTop =480
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13740
                            Top =480
                            Width =1200
                            Height =240
                            Name ="Label91"
                            Caption ="SubFiberCount"
                            LayoutCachedLeft =13740
                            LayoutCachedTop =480
                            LayoutCachedWidth =14940
                            LayoutCachedHeight =720
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    Left =18780
                    Top =5340
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label92"
                    Caption ="Location"
                    LayoutCachedLeft =18780
                    LayoutCachedTop =5340
                    LayoutCachedWidth =19620
                    LayoutCachedHeight =5640
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15420
                    Top =2400
                    Width =2520
                    ColumnOrder =1
                    TabIndex =29
                    Name ="CableExt"

                    LayoutCachedLeft =15420
                    LayoutCachedTop =2400
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =2640
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14100
                            Top =2400
                            Width =900
                            Height =240
                            Name ="Label97"
                            Caption ="CableExt:"
                            LayoutCachedLeft =14100
                            LayoutCachedTop =2400
                            LayoutCachedWidth =15000
                            LayoutCachedHeight =2640
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =15840
                    Top =5340
                    Width =1155
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label98"
                    Caption ="Cable Rating"
                    LayoutCachedLeft =15840
                    LayoutCachedTop =5340
                    LayoutCachedWidth =16995
                    LayoutCachedHeight =5640
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9120
                    Top =1680
                    Width =3660
                    Height =300
                    ColumnOrder =15
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="JacketColor"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeJacketColor"
                    ColumnWidths ="0;2520;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =1980
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7440
                            Top =1680
                            Width =1065
                            Height =300
                            Name ="txtJacketColor"
                            Caption ="Jacket Color"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =1680
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =1980
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =19620
                    Top =5340
                    Width =480
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label100"
                    Caption ="Jacket"
                    LayoutCachedLeft =19620
                    LayoutCachedTop =5340
                    LayoutCachedWidth =20100
                    LayoutCachedHeight =5640
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =1200
                    Width =3660
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =30
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"10\";\"60\""
                    Name ="CableRating"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeCableRating"
                    ColumnWidths ="1440;1440;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1680
                            Top =1200
                            Width =1065
                            Height =300
                            Name ="Label103"
                            Caption ="Cable Rating"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2745
                            LayoutCachedHeight =1500
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =2340
                    Top =5340
                    Width =1230
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label104"
                    Caption ="Fiber Bundles"
                    LayoutCachedLeft =2340
                    LayoutCachedTop =5340
                    LayoutCachedWidth =3570
                    LayoutCachedHeight =5640
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
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =11280
                    Top =60
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="Text0"
                    ControlSource ="Fiber"

                    LayoutCachedLeft =11280
                    LayoutCachedTop =60
                    LayoutCachedWidth =12960
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
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="strBase"
                    ControlSource ="Base"

                    LayoutCachedLeft =14160
                    LayoutCachedTop =60
                    LayoutCachedWidth =15780
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =11280
                    Top =60
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text16"
                    ControlSource ="Jacket"

                    LayoutCachedLeft =11280
                    LayoutCachedTop =60
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =900
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="Text18"
                    ControlSource ="SubFiber"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =2220
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
                    Left =4920
                    Top =60
                    Width =1020
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txFiberCount"
                    ControlSource ="FiberCount"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =60
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13020
                    Top =60
                    Width =1020
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="Text34"
                    ControlSource ="OD"

                    LayoutCachedLeft =13020
                    LayoutCachedTop =60
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3660
                    Top =60
                    Width =1140
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="Text40"
                    ControlSource ="NumPositions"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =60
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6480
                    Top =60
                    Width =4680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="txDesc"
                    ControlSource ="SalesForceDescription"

                    LayoutCachedLeft =6480
                    LayoutCachedTop =60
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17040
                    Top =60
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="Text44"
                    ControlSource ="TBType"

                    LayoutCachedLeft =17040
                    LayoutCachedTop =60
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =960
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="Text51"
                    ControlSource ="SubOD"

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18780
                    Top =60
                    Width =555
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="CablePasses"
                    ControlSource ="CablePasses"

                    LayoutCachedLeft =18780
                    LayoutCachedTop =60
                    LayoutCachedWidth =19335
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15840
                    Top =60
                    Width =1140
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="txJacket"
                    ControlSource ="JacketType"

                    LayoutCachedLeft =15840
                    LayoutCachedTop =60
                    LayoutCachedWidth =16980
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6000
                    Width =426
                    Height =426
                    TabIndex =12
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

                    LayoutCachedLeft =6000
                    LayoutCachedWidth =6426
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19620
                    Top =60
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="Text101"
                    ControlSource ="CableLevel1"

                    LayoutCachedLeft =19620
                    LayoutCachedTop =60
                    LayoutCachedWidth =19980
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =14
                    Name ="Text105"
                    ControlSource ="FibersPerBundle"

                    LayoutCachedLeft =2280
                    LayoutCachedTop =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =-2147483633
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
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

Sub SearchFilter()
    'Applies filter to critcal fields when searching

    Dim strFilter As String

    strFilter = "[DesignTypeGroupID] like """ & Me.CableType & "*"""
    strFilter = strFilter & " AND SalesForceDescription like """ & Me.Jacket & "*"""
    strFilter = strFilter & " AND FiberCount like """ & Me.FiberCount & "*"""
    strFilter = strFilter & " AND [JacketType] like """ & Me.CableRating & "*"""
    strFilter = strFilter & " AND [Base]  like """ & Me.Base & "*"""

    Me.Filter = strFilter
    Me.FilterOn = True


End Sub

Private Sub RequeryFields()
    'Assign all affected fields to null

    Me.Fiber1 = Null
    Me.Fiber2 = Null
    Me.TBColor = Null
    Me.OD = Null
    Me.Base = Null
    Me.JacketMaterial = Null
    Me.BaseOD = Null
    Me.Description = Null
    Me.TB = Null
    Me.Fiber1 = Null
    Me.Fiber2 = Null
    Me.Fiber1Type = Null
    Me.Fiber2Type = Null
    Me.TBType = Null
    Me.JacketColor = Null
    Me.TBColor = Null
    Me.PrintType = Null
    Me.Fiber2.Visible = False
    Me.Fiber2Type.Visible = False


    Call SearchFilter

End Sub

Private Sub SubColor_Exceptions()
    On Error GoTo ErrorHandler:

    'Special cases of default attributes for cables


    Select Case True
    Case Me.Base Like "pp144*"
        Me.TBColor = "0"
    Case Me.Base Like "pp096*"
        Me.TBColor = "0"
    Case Me.CableType Like "B"
        Me.TBColor = "X"
    Case Me.Location = 2 And Fiber2Type.Visible = True And Me.Mode1 <> Me.Mode2
        Me.TBColor = "X"
    End Select



ErrorHandler:
    Resume Next


End Sub

Private Sub btnFind_Click()

    'Opens construction instance for user
    Call OpenAClient(3, Me.strBase, Nz(Me.CablePasses, 2))
End Sub

Private Sub btnSelect_Click()
    Me.OD = Me.strBase
    Call OD_AfterUpdate
End Sub

Private Sub btnSubmit_Click()
    On Error GoTo Err_OD_AfterUpdate
    
    Dim ItemNo As String
    Dim Armor As String
    Dim Ext   As String
    Dim PrintType As String
    Dim Fiber1Type As String
    Dim Fiber2Type As String
    
    
    'Identify if the print type is "AFL Generic"
    If Me.PrintType = "2A" Then
        PrintType = "2"
    Else
        PrintType = Me.PrintType
    End If
    
    If IsNull(Me.Fiber1Type.Column(3)) = False And Me.Fiber1Type <> "" And Me.Fiber1Type.Column(3) <> "" Then
        Ext = Me.Fiber1Type.Column(3)
    End If
    
    
    Fiber1Type = Me.Fiber1Type.Column(0)
    If IsNull(Me.Fiber2Type.Column(0)) = False Then
        Fiber2Type = Me.Fiber2Type.Column(0)
    End If
    
    If IsNull(Me.CableType) = True Or IsNull(Me.Jacket) = True Or IsNull(Me.FiberCount) = True Or IsNull(Me.OD) = True Or IsNull(Me.Fiber1) = True Or IsNull(Me.Fiber1Type) = True _
                                                                                                                                                                              Or IsNull(Me.TBType) = True Or IsNull(Me.JacketColor) = True Or IsNull(Me.TBColor) = True Or IsNull(Me.PrintType) = True Or Int(Me.Fiber1) + Int(Me.Fiber2) <> Int(Me.FiberCount) Then
    
        MsgBox ("Please fill out the form completely or check fiber counts.")
        Exit Sub
    End If
    

    
    If Me.Fiber2.Visible = False Then
        
        ItemNo = Left(Me.ItemNoPrefix, 5) & Fiber1Type & Mid(Me.ItemNoPrefix, 6, 2) & PrintType & Me.JacketColor & Me.TBColor & Me.TBType.Column(0) & Mid(Me.ItemNoPrefix, 8, 5)
        
    ElseIf Me.Fiber2.Visible = True Then
        
        ItemNo = Left(Me.ItemNoPrefix, 5) & "X" & Mid(Me.ItemNoPrefix, 6, 2) & PrintType & Me.JacketColor & Me.TBColor & Me.TBType.Column(0) & Mid(Me.ItemNoPrefix, 8, 5) & "-" _
                                                                                                                                                                        & Me.Fiber1 & Me.Mode1 & Fiber1Type & "/" & Me.Fiber2 & Me.Mode2 & Fiber2Type
        
    End If
    
    If Me.Armor = -1 Then
        Select Case Me.CableRating
        Case "P", "Q"
            Armor = "-AIAP"
        Case "R"
            Armor = "-AIAR"
        Case "E"
            Armor = "-AIAL"
        Case Else
            Armor = ""
        End Select
    Else
        Armor = ""
    End If
    
    If Me.CableExt = "NONE" Then
        Me.CableExt = Null
    End If
    
    If IsNull(Me.CableExt) = False And Ext <> "" Then
        Ext = Me.CableExt & Ext
    ElseIf IsNull(Me.CableExt) = False And Ext = "" Then
        Ext = "-" & Me.CableExt
    ElseIf IsNull(Me.CableExt) = True And Ext <> "" Then
        Ext = Ext
    End If

    
    
    Me.ItemNo = ItemNo & Ext & Armor
    
    MsgBox "Design Code : " & Me.ItemNo
    
    




Exit_OD_AfterUpdate:
    DoCmd.Close acForm, "frmDesignCodeOD"


    Exit Sub


Err_OD_AfterUpdate:
    MsgBox Err.Number & " " & Err.Description
    DoCmd.Close acForm, "frmDesignCodeOD"

End Sub

Private Sub CableRating_AfterUpdate()
    Me.FiberCount = Null
    Me.FiberCount.Requery
    Me.OD = Null
    Me.OD.Requery
    Me.Base = Null

    Call SearchFilter

End Sub

Private Sub Fiber1Type_AfterUpdate()


    Me.Mode1 = Me.Fiber1Type.Column(2)

    Me.JacketColor = Me.Fiber1Type.Column(4)



    If Location = 2 Then
        Call JacketColor_AfterUpdate
    End If

    If Me.Fiber1Type = "X" Then
    
    End If


End Sub

Private Sub Fiber2Type_AfterUpdate()
    'Select Case Me.Fiber2Type
    'Case "F", "5", "6", "7", "8", "A", "B", "C", "D", "E", "1", "L", "N", "O", "Y", "CRG", "CRG+"
    '   Me.Mode2 = "MM"
    'Case Else
    '   Me.Mode2 = "SM"
    'End Select

    Me.Mode2 = Me.Fiber2Type.Column(2)

    If Me.Fiber2.Visible = True And Me.Mode1 = "SM" Then
        MsgBox ("Singlemode fiber should be Fiber 2")
    End If
End Sub

Private Sub Fiber1_AfterUpdate()
    Me.Fiber2 = Null
    Me.Fiber2.Requery

    Dim twips As Integer

    twips = 1440

    If Int(Me.Fiber1) < Int(Me.FiberCount) Then
        Me.Fiber2.Visible = True
        Me.Fiber2Type.Visible = True
        Me.txtFiber2.Visible = True
        Me.txtFiber2Type.Visible = True
        Me.TBType.top = 1.4583 * twips
        Me.JacketColor.top = 1.75 * twips
        Me.TBColor.top = 2.0417 * twips
        Me.PrintType.top = 2.3333 * twips
        Me.txtTBType.top = 1.4583 * twips
        Me.txtJacketColor.top = 1.75 * twips
        Me.txtTBColor.top = 2.0417 * twips
        Me.txtPrintType.top = 2.3333 * twips
    Else
        Me.Fiber2.Visible = False
        Me.Fiber2Type.Visible = False
        Me.txtFiber2.Visible = False
        Me.txtFiber2Type.Visible = False
        Me.TBType.top = 0.875 * twips
        Me.JacketColor.top = 1.1667 * twips
        Me.TBColor.top = 1.4583 * twips
        Me.PrintType.top = 1.75 * twips
        Me.txtTBType.top = 0.875 * twips
        Me.txtJacketColor.top = 1.1667 * twips
        Me.txtTBColor.top = 1.4583 * twips
        Me.txtPrintType.top = 1.75 * twips
    
    End If
End Sub

Private Sub FiberCount_AfterUpdate()
    Call RequeryFields
    Me.OD.Requery

    Me.Fiber1 = Me.FiberCount

    Call Fiber1_AfterUpdate

End Sub

Private Sub Cable_AfterUpdate()

    Me.Location = Me.Cable.Column(2)
    Me.CableType = Me.Cable.Column(0)
    Call RequeryFields
    Me.Jacket = Null
    Me.FiberCount = Null

    Me.Jacket.Requery
    Me.FiberCount.Requery
    Me.CableRating = Null
    Me.CableRating.Requery

    Call Fiber1_AfterUpdate

    If Me.Location = 2 Then
        Me.OD.ColumnCount = 8
    Else
        Me.OD.ColumnCount = 8
    End If
    Call SearchFilter

End Sub

Private Sub Jacket_AfterUpdate()


    Me.FiberCount = Null
    Me.FiberCount.Requery
    Me.OD = Null
    Me.OD.Requery
    Me.CableRating = Null
    Me.CableRating.Requery
    Call RequeryFields

End Sub

Private Sub JacketColor_AfterUpdate()
    If Me.Location = 2 Then
        Me.TBColor = Me.JacketColor
        Me.TBColor.Requery
        Call SubColor_Exceptions

    ElseIf Me.Location = 1 Then
        If Me.Jacket = "F" Then
    
        Else
            Me.TBColor = "0"
            Me.TBColor.Requery
        End If
    
    
    End If
End Sub

Private Sub OD_AfterUpdate()
    On Error GoTo Err_OD_AfterUpdate

    Me.Base = Me.OD.Column(0)

    Call SearchFilter


    DoCmd.OpenForm "frmDesignCodeOD"



    Me.JacketMaterial = Forms!frmDesignCodeOD!JacketMaterial
    Me.Location = Forms!frmDesignCodeOD!CablePasses
    Me.BaseOD = Forms!frmDesignCodeOD!NominalOD
    Me.Description = Me.OD.Column(5)
    Me.TB = Forms!frmDesignCodeOD!TBType
    Me.ItemNoPrefix = Forms!frmDesignCodeOD!DesignCode
    Me.Jacket = Me.txDesc
    Me.CableRating.Requery
    Me.CableRating = Me.txJacket
    Me.SubFiberCount = Int(Me.OD.Column(3))
    Me.CableExt = Forms!frmDesignCodeOD!DesignExtension



    DoCmd.Close acForm, "frmDesignCodeOD"





    'Checking if dual or single pass cable and adjusting
    If Me.Location = 1 Then

        Me.TBColor.rowSource = "qryDesignCodeTBColorFiltered"
        Me.TBColor = 0
        Me.Fiber1.rowSource = "qryDesignCodeFiber1Count1Pass"
        Me.Fiber2.rowSource = "qryDesignCodeFiber2Count1Pass"
    
    Else

        Me.TBColor.rowSource = "qryDesignCodeTBColor"
        Me.Fiber1.rowSource = "qryDesignCodeFiber1Count2Pass"
        Me.Fiber2.rowSource = "qryDesignCodeFiber2Count2Pass"
    End If
    Me.TBColor.Requery

    ' For furcation tubes
    If Me.TB = "Furcation" Then
        'Me.ItemNoPrefix = Me.CableType & "F" & "000" & Mid(Me.ItemNoPrefix, 6, 20)
        Me.TBType = 0
        Me.TBColor.rowSource = "qryDesignCodeFurcationTube"
        Me.TBColor.Locked = False
        Me.txtTBColor.Caption = "Furc Tube"
        Me.TBColor = "P"
        Me.Fiber1Type.rowSource = "qryDesignCodeFiberTypeFurcation"
        Me.Fiber1Type.value = "@"
        Me.Fiber1Type = Me.Fiber1Type.ItemData(1)
    

    Else
        Me.Fiber1Type.rowSource = "qryDesignCodeFiberType"
        Me.txtTBColor.Caption = "Sub Color"
    End If



    Me.FiberCount.Requery
    Me.FiberCount = Me.OD.Column(1)
    Me.Fiber1.Requery
    Me.Fiber1 = FiberCount
    Me.OD.Requery
    Me.TBType.Requery
    Me.TBType = Me.TBType.ItemData(0)



    Call Fiber1_AfterUpdate
    Call SearchFilter

Exit_OD_AfterUpdate:
    DoCmd.Close acForm, "frmDesignCodeOD"


    Exit Sub


Err_OD_AfterUpdate:
    DoCmd.Close acForm, "frmDesignCodeOD"
    MsgBox Err.Description
    Resume Exit_OD_AfterUpdate




End Sub

Private Sub Refresh_Click()
    Call SearchFilter
End Sub

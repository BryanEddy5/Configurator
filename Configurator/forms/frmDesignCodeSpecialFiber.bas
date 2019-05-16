Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    TabularFamily =89
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7320
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =6750
    Top =2160
    Right =13725
    Bottom =4710
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0746e86bf49be440
    End
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =3000
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    ListWidth =4320
                    Left =2100
                    Top =840
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Vendor"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeVendor"
                    ColumnWidths ="1440;1440"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =840
                            Width =1065
                            Height =300
                            Name ="Label3"
                            Caption ="Vendor"
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
                    ListRows =12
                    ListWidth =4320
                    Left =2100
                    Top =1260
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="Fiber1Type"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeFiberType"
                    ColumnWidths ="0;1800;0"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =1260
                            Width =1065
                            Height =300
                            Name ="Label38"
                            Caption ="Diameter"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ListRows =12
                    ListWidth =4320
                    Left =2100
                    Top =420
                    Width =3660
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="ITU"
                    RowSourceType ="Table/Query"
                    RowSource ="qryDesignCodeITU"
                    ColumnWidths ="1440;1800;0"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =420
                            Width =1065
                            Height =300
                            Name ="Label1"
                            Caption ="ITU"
                        End
                    End
                End
            End
        End
    End
End

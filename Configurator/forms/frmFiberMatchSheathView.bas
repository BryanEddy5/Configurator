Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =30420
    DatasheetFontHeight =10
    ItemSuffix =182
    Top =2835
    Right =19050
    Bottom =6930
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb941ee872b99e440
    End
    RecordSource ="qryfrmFiberMatchExistingFindSheath"
    Caption ="qryfrmFiberMatchExistingFindSheath"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4110
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =120
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    Name ="New Oracle Part Number"
                    ControlSource ="New Oracle Part Number"
                    EventProcPrefix ="New_Oracle_Part_Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =720
                            Height =255
                            Name ="New Oracle Part Number_Label"
                            Caption ="New Oracle Part Number"
                            EventProcPrefix ="New_Oracle_Part_Number_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =660
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Item No"
                    ControlSource ="Item No"
                    EventProcPrefix ="Item_No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =660
                            Width =720
                            Height =255
                            Name ="Item No_Label"
                            Caption ="Item No"
                            EventProcPrefix ="Item_No_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =1020
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Customer"
                    ControlSource ="Customer"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1020
                            Width =720
                            Height =255
                            Name ="Customer_Label"
                            Caption ="Customer"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =1560
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =3
                    Name ="Active"
                    ControlSource ="Active"
                    Format ="Yes/No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =720
                            Height =255
                            Name ="Active_Label"
                            Caption ="Active"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =1920
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="Reason"
                    ControlSource ="Reason"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =720
                            Height =255
                            Name ="Reason_Label"
                            Caption ="Reason"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =2460
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="Base"
                    ControlSource ="Base"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2460
                            Width =720
                            Height =255
                            Name ="Base_Label"
                            Caption ="Base"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =2820
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =6
                    Name ="Revision Date"
                    ControlSource ="Revision Date"
                    EventProcPrefix ="Revision_Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2820
                            Width =720
                            Height =255
                            Name ="Revision Date_Label"
                            Caption ="Revision Date"
                            EventProcPrefix ="Revision_Date_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =3180
                    Width =390
                    Height =255
                    ColumnWidth =390
                    TabIndex =7
                    Name ="Revision Letter"
                    ControlSource ="Revision Letter"
                    EventProcPrefix ="Revision_Letter"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3180
                            Width =720
                            Height =255
                            Name ="Revision Letter_Label"
                            Caption ="Revision Letter"
                            EventProcPrefix ="Revision_Letter_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =120
                    Width =1920
                    Height =840
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="Revision History"
                    ControlSource ="Revision History"
                    EventProcPrefix ="Revision_History"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2820
                            Top =120
                            Width =720
                            Height =255
                            Name ="Revision History_Label"
                            Caption ="Revision History"
                            EventProcPrefix ="Revision_History_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =1080
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =9
                    Name ="PSS Document #"
                    ControlSource ="PSS Document #"
                    EventProcPrefix ="PSS_Document__"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2820
                            Top =1080
                            Width =720
                            Height =255
                            Name ="PSS Document #_Label"
                            Caption ="PSS Document #"
                            EventProcPrefix ="PSS_Document___Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =1620
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="Customer Part Num"
                    ControlSource ="Customer Part Num"
                    EventProcPrefix ="Customer_Part_Num"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2820
                            Top =1620
                            Width =720
                            Height =255
                            Name ="Customer Part Num_Label"
                            Caption ="Customer Part Num"
                            EventProcPrefix ="Customer_Part_Num_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =2160
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =11
                    Name ="Print Item No"
                    ControlSource ="Print Item No"
                    Format ="Yes/No"
                    EventProcPrefix ="Print_Item_No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2820
                            Top =2160
                            Width =720
                            Height =255
                            Name ="Print Item No_Label"
                            Caption ="Print Item No"
                            EventProcPrefix ="Print_Item_No_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =2520
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =12
                    Name ="Print Reel No"
                    ControlSource ="Print Reel No"
                    Format ="Yes/No"
                    EventProcPrefix ="Print_Reel_No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2820
                            Top =2520
                            Width =720
                            Height =255
                            Name ="Print Reel No_Label"
                            Caption ="Print Reel No"
                            EventProcPrefix ="Print_Reel_No_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =2880
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =13
                    Name ="Print Height"
                    ControlSource ="Print Height"
                    EventProcPrefix ="Print_Height"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2820
                            Top =2880
                            Width =720
                            Height =255
                            Name ="Print Height_Label"
                            Caption ="Print Height"
                            EventProcPrefix ="Print_Height_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =3420
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =14
                    Name ="Print Type (base)"
                    ControlSource ="Print Type (base)"
                    EventProcPrefix ="Print_Type__base_"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2820
                            Top =3420
                            Width =720
                            Height =255
                            Name ="Print Type (base)_Label"
                            Caption ="Print Type (base)"
                            EventProcPrefix ="Print_Type__base__Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =120
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =15
                    Name ="Print Line 1"
                    ControlSource ="Print Line 1"
                    EventProcPrefix ="Print_Line_1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =120
                            Width =720
                            Height =255
                            Name ="Print Line 1_Label"
                            Caption ="Print Line 1"
                            EventProcPrefix ="Print_Line_1_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =660
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =16
                    Name ="Print Line 2"
                    ControlSource ="Print Line 2"
                    EventProcPrefix ="Print_Line_2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =660
                            Width =720
                            Height =255
                            Name ="Print Line 2_Label"
                            Caption ="Print Line 2"
                            EventProcPrefix ="Print_Line_2_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =1200
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =17
                    Name ="Print Line 3"
                    ControlSource ="Print Line 3"
                    EventProcPrefix ="Print_Line_3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =1200
                            Width =720
                            Height =255
                            Name ="Print Line 3_Label"
                            Caption ="Print Line 3"
                            EventProcPrefix ="Print_Line_3_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =1740
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =18
                    Name ="Print Line 4"
                    ControlSource ="Print Line 4"
                    EventProcPrefix ="Print_Line_4"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =1740
                            Width =720
                            Height =255
                            Name ="Print Line 4_Label"
                            Caption ="Print Line 4"
                            EventProcPrefix ="Print_Line_4_Label"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =15
                    ListWidth =5184
                    Left =6360
                    Top =2280
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="NewPrintLine4"
                    ControlSource ="NewPrintLine4"
                    RowSourceType ="Table/Query"
                    RowSource ="tblLine4PrintCode"
                    ColumnWidths ="14;4032;1008"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =2280
                            Width =720
                            Height =255
                            Name ="NewPrintLine4_Label"
                            Caption ="NewPrintLine4"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =2640
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =20
                    Name ="PrintSpacing"
                    ControlSource ="PrintSpacing"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =2640
                            Width =720
                            Height =255
                            Name ="PrintSpacing_Label"
                            Caption ="PrintSpacing"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =3180
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =21
                    Name ="PrintNotes"
                    ControlSource ="PrintNotes"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =3180
                            Width =720
                            Height =255
                            Name ="PrintNotes_Label"
                            Caption ="PrintNotes"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =3720
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =22
                    Name ="Listing Company"
                    ControlSource ="Listing Company"
                    EventProcPrefix ="Listing_Company"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =3720
                            Width =720
                            Height =255
                            Name ="Listing Company_Label"
                            Caption ="Listing Company"
                            EventProcPrefix ="Listing_Company_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =120
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =23
                    Name ="UL-ETL-Listing"
                    ControlSource ="UL-ETL-Listing"
                    EventProcPrefix ="UL_ETL_Listing"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =120
                            Width =720
                            Height =255
                            Name ="UL-ETL-Listing_Label"
                            Caption ="UL-ETL-Listing"
                            EventProcPrefix ="UL_ETL_Listing_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =480
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =24
                    Name ="UL-ETL-Const"
                    ControlSource ="UL-ETL-Const"
                    EventProcPrefix ="UL_ETL_Const"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =480
                            Width =720
                            Height =255
                            Name ="UL-ETL-Const_Label"
                            Caption ="UL-ETL-Const"
                            EventProcPrefix ="UL_ETL_Const_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =840
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =25
                    Name ="UL-ETL-Section"
                    ControlSource ="UL-ETL-Section"
                    EventProcPrefix ="UL_ETL_Section"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =840
                            Width =720
                            Height =255
                            Name ="UL-ETL-Section_Label"
                            Caption ="UL-ETL-Section"
                            EventProcPrefix ="UL_ETL_Section_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =1200
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =26
                    Name ="Unit Series"
                    ControlSource ="Unit Series"
                    EventProcPrefix ="Unit_Series"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =1200
                            Width =720
                            Height =255
                            Name ="Unit Series_Label"
                            Caption ="Unit Series"
                            EventProcPrefix ="Unit_Series_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =1560
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =27
                    Name ="FRP Dia"
                    ControlSource ="FRP Dia"
                    EventProcPrefix ="FRP_Dia"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =1560
                            Width =720
                            Height =255
                            Name ="FRP Dia_Label"
                            Caption ="FRP Dia"
                            EventProcPrefix ="FRP_Dia_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =1920
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =28
                    Name ="CM Type"
                    ControlSource ="CM Type"
                    EventProcPrefix ="CM_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =1920
                            Width =720
                            Height =255
                            Name ="CM Type_Label"
                            Caption ="CM Type"
                            EventProcPrefix ="CM_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =2460
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =29
                    Name ="CM Matrl"
                    ControlSource ="CM Matrl"
                    EventProcPrefix ="CM_Matrl"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =2460
                            Width =720
                            Height =255
                            Name ="CM Matrl_Label"
                            Caption ="CM Matrl"
                            EventProcPrefix ="CM_Matrl_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =3000
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =30
                    Name ="CM OD"
                    ControlSource ="CM OD"
                    EventProcPrefix ="CM_OD"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =3000
                            Width =720
                            Height =255
                            Name ="CM OD_Label"
                            Caption ="CM OD"
                            EventProcPrefix ="CM_OD_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =3360
                    Width =1560
                    Height =255
                    ColumnWidth =1560
                    TabIndex =31
                    Name ="CM OD Tol +"
                    ControlSource ="CM OD Tol +"
                    EventProcPrefix ="CM_OD_Tol__"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =3360
                            Width =720
                            Height =255
                            Name ="CM OD Tol +_Label"
                            Caption ="CM OD Tol +"
                            EventProcPrefix ="CM_OD_Tol___Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =3720
                    Width =1560
                    Height =255
                    ColumnWidth =1560
                    TabIndex =32
                    Name ="CM OD Tol -6"
                    ControlSource ="CM OD Tol -"
                    EventProcPrefix ="CM_OD_Tol__6"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =3720
                            Width =720
                            Height =255
                            Name ="CM OD Tol -6_Label"
                            Caption ="CM OD Tol -6"
                            EventProcPrefix ="CM_OD_Tol__6_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =120
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =33
                    Name ="CM Wall"
                    ControlSource ="CM Wall"
                    EventProcPrefix ="CM_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =120
                            Width =720
                            Height =255
                            Name ="CM Wall_Label"
                            Caption ="CM Wall"
                            EventProcPrefix ="CM_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =480
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =34
                    Name ="CM Min Wall"
                    ControlSource ="CM Min Wall"
                    EventProcPrefix ="CM_Min_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =480
                            Width =720
                            Height =255
                            Name ="CM Min Wall_Label"
                            Caption ="CM Min Wall"
                            EventProcPrefix ="CM_Min_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =840
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =35
                    Name ="Binder #1"
                    ControlSource ="Binder #1"
                    EventProcPrefix ="Binder__1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =840
                            Width =720
                            Height =255
                            Name ="Binder #1_Label"
                            Caption ="Binder #1"
                            EventProcPrefix ="Binder__1_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =1380
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =36
                    Name ="Qty Binder #1"
                    ControlSource ="Qty Binder #1"
                    EventProcPrefix ="Qty_Binder__1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =1380
                            Width =720
                            Height =255
                            Name ="Qty Binder #1_Label"
                            Caption ="Qty Binder #1"
                            EventProcPrefix ="Qty_Binder__1_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =1740
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =37
                    Name ="Binder #2"
                    ControlSource ="Binder #2"
                    EventProcPrefix ="Binder__2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =1740
                            Width =720
                            Height =255
                            Name ="Binder #2_Label"
                            Caption ="Binder #2"
                            EventProcPrefix ="Binder__2_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =2280
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =38
                    Name ="Qty Binder #2"
                    ControlSource ="Qty Binder #2"
                    EventProcPrefix ="Qty_Binder__2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =2280
                            Width =720
                            Height =255
                            Name ="Qty Binder #2_Label"
                            Caption ="Qty Binder #2"
                            EventProcPrefix ="Qty_Binder__2_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =2640
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =39
                    Name ="Lay Length"
                    ControlSource ="Lay Length"
                    EventProcPrefix ="Lay_Length"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =2640
                            Width =720
                            Height =255
                            Name ="Lay Length_Label"
                            Caption ="Lay Length"
                            EventProcPrefix ="Lay_Length_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =3180
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =40
                    Name ="Helix Factor"
                    ControlSource ="Helix Factor"
                    EventProcPrefix ="Helix_Factor"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =3180
                            Width =720
                            Height =255
                            Name ="Helix Factor_Label"
                            Caption ="Helix Factor"
                            EventProcPrefix ="Helix_Factor_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =3540
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =41
                    Name ="Core Wrap"
                    ControlSource ="Core Wrap"
                    EventProcPrefix ="Core_Wrap"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11100
                            Top =3540
                            Width =720
                            Height =255
                            Name ="Core Wrap_Label"
                            Caption ="Core Wrap"
                            EventProcPrefix ="Core_Wrap_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =120
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =42
                    Name ="Binder #3"
                    ControlSource ="Binder #3"
                    EventProcPrefix ="Binder__3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =120
                            Width =720
                            Height =255
                            Name ="Binder #3_Label"
                            Caption ="Binder #3"
                            EventProcPrefix ="Binder__3_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =660
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =43
                    Name ="Qty Binder #3"
                    ControlSource ="Qty Binder #3"
                    EventProcPrefix ="Qty_Binder__3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =660
                            Width =720
                            Height =255
                            Name ="Qty Binder #3_Label"
                            Caption ="Qty Binder #3"
                            EventProcPrefix ="Qty_Binder__3_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =1020
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =44
                    Name ="Core Dia"
                    ControlSource ="Core Dia"
                    Format ="Fixed"
                    EventProcPrefix ="Core_Dia"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =1020
                            Width =720
                            Height =255
                            Name ="Core Dia_Label"
                            Caption ="Core Dia"
                            EventProcPrefix ="Core_Dia_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =1380
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =45
                    Name ="Ripcord"
                    ControlSource ="Ripcord"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =1380
                            Width =720
                            Height =255
                            Name ="Ripcord_Label"
                            Caption ="Ripcord"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =1920
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =46
                    Name ="Jacket Material"
                    ControlSource ="Jacket Material"
                    EventProcPrefix ="Jacket_Material"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =1920
                            Width =720
                            Height =255
                            Name ="Jacket Material_Label"
                            Caption ="Jacket Material"
                            EventProcPrefix ="Jacket_Material_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =2460
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =47
                    Name ="Jacket Color"
                    ControlSource ="Jacket Color"
                    EventProcPrefix ="Jacket_Color"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =2460
                            Width =720
                            Height =255
                            Name ="Jacket Color_Label"
                            Caption ="Jacket Color"
                            EventProcPrefix ="Jacket_Color_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =3000
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =48
                    Name ="Nom Wall"
                    ControlSource ="Nom Wall"
                    EventProcPrefix ="Nom_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =3000
                            Width =720
                            Height =255
                            Name ="Nom Wall_Label"
                            Caption ="Nom Wall"
                            EventProcPrefix ="Nom_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =3360
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =49
                    Name ="Min Ave Wall"
                    ControlSource ="Min Ave Wall"
                    EventProcPrefix ="Min_Ave_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =3360
                            Width =720
                            Height =255
                            Name ="Min Ave Wall_Label"
                            Caption ="Min Ave Wall"
                            EventProcPrefix ="Min_Ave_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14640
                    Top =3720
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =50
                    Name ="Min Spot Wall"
                    ControlSource ="Min Spot Wall"
                    EventProcPrefix ="Min_Spot_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13860
                            Top =3720
                            Width =720
                            Height =255
                            Name ="Min Spot Wall_Label"
                            Caption ="Min Spot Wall"
                            EventProcPrefix ="Min_Spot_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =120
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =51
                    Name ="Max Ave Wall"
                    ControlSource ="Max Ave Wall"
                    EventProcPrefix ="Max_Ave_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =120
                            Width =720
                            Height =255
                            Name ="Max Ave Wall_Label"
                            Caption ="Max Ave Wall"
                            EventProcPrefix ="Max_Ave_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =480
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =52
                    Name ="Nominal OD"
                    ControlSource ="Nominal OD"
                    EventProcPrefix ="Nominal_OD"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =480
                            Width =720
                            Height =255
                            Name ="Nominal OD_Label"
                            Caption ="Nominal OD"
                            EventProcPrefix ="Nominal_OD_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =840
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =53
                    Name ="OD Tol  (+)"
                    ControlSource ="OD Tol  (+)"
                    EventProcPrefix ="OD_Tol_____"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =840
                            Width =720
                            Height =255
                            Name ="OD Tol  (+)_Label"
                            Caption ="OD Tol  (+)"
                            EventProcPrefix ="OD_Tol______Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =1200
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =54
                    Name ="OD Tol  (-)7"
                    ControlSource ="OD Tol  (-)"
                    EventProcPrefix ="OD_Tol_____7"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =1200
                            Width =720
                            Height =255
                            Name ="OD Tol  (-)7_Label"
                            Caption ="OD Tol  (-)7"
                            EventProcPrefix ="OD_Tol_____7_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =1560
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =55
                    Name ="Chip Type"
                    ControlSource ="Chip Type"
                    EventProcPrefix ="Chip_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =1560
                            Width =720
                            Height =255
                            Name ="Chip Type_Label"
                            Caption ="Chip Type"
                            EventProcPrefix ="Chip_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =1920
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =56
                    Name ="Fiber Type A"
                    ControlSource ="Fiber Type A"
                    EventProcPrefix ="Fiber_Type_A"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =1920
                            Width =720
                            Height =255
                            Name ="Fiber Type A_Label"
                            Caption ="Fiber Type A"
                            EventProcPrefix ="Fiber_Type_A_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =2460
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =57
                    Name ="1st Req Freq - A"
                    ControlSource ="1st Req Freq - A"
                    EventProcPrefix ="Ctl1st_Req_Freq___A"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =2460
                            Width =720
                            Height =255
                            Name ="1st Req Freq - A_Label"
                            Caption ="1st Req Freq - A"
                            EventProcPrefix ="Ctl1st_Req_Freq___A_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =2820
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =58
                    Name ="1st Max Atten - A"
                    ControlSource ="1st Max Atten - A"
                    EventProcPrefix ="Ctl1st_Max_Atten___A"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =2820
                            Width =720
                            Height =255
                            Name ="1st Max Atten - A_Label"
                            Caption ="1st Max Atten - A"
                            EventProcPrefix ="Ctl1st_Max_Atten___A_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =3180
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =59
                    Name ="1st Min BandW - A"
                    ControlSource ="1st Min BandW - A"
                    EventProcPrefix ="Ctl1st_Min_BandW___A"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =3180
                            Width =720
                            Height =255
                            Name ="1st Min BandW - A_Label"
                            Caption ="1st Min BandW - A"
                            EventProcPrefix ="Ctl1st_Min_BandW___A_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17400
                    Top =3540
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =60
                    Name ="2nd Req Freq - A"
                    ControlSource ="2nd Req Freq - A"
                    EventProcPrefix ="Ctl2nd_Req_Freq___A"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16620
                            Top =3540
                            Width =720
                            Height =255
                            Name ="2nd Req Freq - A_Label"
                            Caption ="2nd Req Freq - A"
                            EventProcPrefix ="Ctl2nd_Req_Freq___A_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =120
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =61
                    Name ="2nd Max Atten - A"
                    ControlSource ="2nd Max Atten - A"
                    EventProcPrefix ="Ctl2nd_Max_Atten___A"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =120
                            Width =720
                            Height =255
                            Name ="2nd Max Atten - A_Label"
                            Caption ="2nd Max Atten - A"
                            EventProcPrefix ="Ctl2nd_Max_Atten___A_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =480
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =62
                    Name ="2nd Min BandW - A"
                    ControlSource ="2nd Min BandW - A"
                    EventProcPrefix ="Ctl2nd_Min_BandW___A"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =480
                            Width =720
                            Height =255
                            Name ="2nd Min BandW - A_Label"
                            Caption ="2nd Min BandW - A"
                            EventProcPrefix ="Ctl2nd_Min_BandW___A_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =840
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =63
                    Name ="Fiber Type B"
                    ControlSource ="Fiber Type B"
                    EventProcPrefix ="Fiber_Type_B"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =840
                            Width =720
                            Height =255
                            Name ="Fiber Type B_Label"
                            Caption ="Fiber Type B"
                            EventProcPrefix ="Fiber_Type_B_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =1380
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =64
                    Name ="1st Req Freq - B"
                    ControlSource ="1st Req Freq - B"
                    EventProcPrefix ="Ctl1st_Req_Freq___B"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =1380
                            Width =720
                            Height =255
                            Name ="1st Req Freq - B_Label"
                            Caption ="1st Req Freq - B"
                            EventProcPrefix ="Ctl1st_Req_Freq___B_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =1740
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =65
                    Name ="1st Max Atten - B"
                    ControlSource ="1st Max Atten - B"
                    EventProcPrefix ="Ctl1st_Max_Atten___B"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =1740
                            Width =720
                            Height =255
                            Name ="1st Max Atten - B_Label"
                            Caption ="1st Max Atten - B"
                            EventProcPrefix ="Ctl1st_Max_Atten___B_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =2100
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =66
                    Name ="1st Min BandW - B"
                    ControlSource ="1st Min BandW - B"
                    EventProcPrefix ="Ctl1st_Min_BandW___B"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =2100
                            Width =720
                            Height =255
                            Name ="1st Min BandW - B_Label"
                            Caption ="1st Min BandW - B"
                            EventProcPrefix ="Ctl1st_Min_BandW___B_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =2460
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =67
                    Name ="2nd Req Freq - B"
                    ControlSource ="2nd Req Freq - B"
                    EventProcPrefix ="Ctl2nd_Req_Freq___B"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =2460
                            Width =720
                            Height =255
                            Name ="2nd Req Freq - B_Label"
                            Caption ="2nd Req Freq - B"
                            EventProcPrefix ="Ctl2nd_Req_Freq___B_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =2820
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =68
                    Name ="2nd Max Atten - B"
                    ControlSource ="2nd Max Atten - B"
                    EventProcPrefix ="Ctl2nd_Max_Atten___B"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =2820
                            Width =720
                            Height =255
                            Name ="2nd Max Atten - B_Label"
                            Caption ="2nd Max Atten - B"
                            EventProcPrefix ="Ctl2nd_Max_Atten___B_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =3180
                    Width =540
                    Height =255
                    ColumnWidth =540
                    TabIndex =69
                    Name ="2nd Min BandW - B"
                    ControlSource ="2nd Min BandW - B"
                    EventProcPrefix ="Ctl2nd_Min_BandW___B"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =3180
                            Width =720
                            Height =255
                            Name ="2nd Min BandW - B_Label"
                            Caption ="2nd Min BandW - B"
                            EventProcPrefix ="Ctl2nd_Min_BandW___B_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20160
                    Top =3540
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =70
                    Name ="Fiber Type C"
                    ControlSource ="Fiber Type C"
                    EventProcPrefix ="Fiber_Type_C"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =19380
                            Top =3540
                            Width =720
                            Height =255
                            Name ="Fiber Type C_Label"
                            Caption ="Fiber Type C"
                            EventProcPrefix ="Fiber_Type_C_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =120
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =71
                    Name ="1st Req Freq - C"
                    ControlSource ="1st Req Freq - C"
                    EventProcPrefix ="Ctl1st_Req_Freq___C"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =120
                            Width =720
                            Height =255
                            Name ="1st Req Freq - C_Label"
                            Caption ="1st Req Freq - C"
                            EventProcPrefix ="Ctl1st_Req_Freq___C_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =480
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =72
                    Name ="1st Max Atten - C"
                    ControlSource ="1st Max Atten - C"
                    EventProcPrefix ="Ctl1st_Max_Atten___C"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =480
                            Width =720
                            Height =255
                            Name ="1st Max Atten - C_Label"
                            Caption ="1st Max Atten - C"
                            EventProcPrefix ="Ctl1st_Max_Atten___C_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =840
                    Width =540
                    Height =255
                    ColumnWidth =540
                    TabIndex =73
                    Name ="1st Min BandW - C"
                    ControlSource ="1st Min BandW - C"
                    EventProcPrefix ="Ctl1st_Min_BandW___C"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =840
                            Width =720
                            Height =255
                            Name ="1st Min BandW - C_Label"
                            Caption ="1st Min BandW - C"
                            EventProcPrefix ="Ctl1st_Min_BandW___C_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =1200
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =74
                    Name ="2nd Req Freq - C"
                    ControlSource ="2nd Req Freq - C"
                    EventProcPrefix ="Ctl2nd_Req_Freq___C"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =1200
                            Width =720
                            Height =255
                            Name ="2nd Req Freq - C_Label"
                            Caption ="2nd Req Freq - C"
                            EventProcPrefix ="Ctl2nd_Req_Freq___C_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =1560
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =75
                    Name ="2nd Max Atten - C"
                    ControlSource ="2nd Max Atten - C"
                    EventProcPrefix ="Ctl2nd_Max_Atten___C"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =1560
                            Width =720
                            Height =255
                            Name ="2nd Max Atten - C_Label"
                            Caption ="2nd Max Atten - C"
                            EventProcPrefix ="Ctl2nd_Max_Atten___C_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =1920
                    Width =540
                    Height =255
                    ColumnWidth =540
                    TabIndex =76
                    Name ="2nd Min BandW - C"
                    ControlSource ="2nd Min BandW - C"
                    EventProcPrefix ="Ctl2nd_Min_BandW___C"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =1920
                            Width =720
                            Height =255
                            Name ="2nd Min BandW - C_Label"
                            Caption ="2nd Min BandW - C"
                            EventProcPrefix ="Ctl2nd_Min_BandW___C_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =2280
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =77
                    Name ="Special Instr Product1"
                    ControlSource ="Special Instr Product1"
                    EventProcPrefix ="Special_Instr_Product1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =2280
                            Width =720
                            Height =255
                            Name ="Special Instr Product1_Label"
                            Caption ="Special Instr Product1"
                            EventProcPrefix ="Special_Instr_Product1_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =2820
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =78
                    Name ="Special Instr Product2"
                    ControlSource ="Special Instr Product2"
                    EventProcPrefix ="Special_Instr_Product2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =2820
                            Width =720
                            Height =255
                            Name ="Special Instr Product2_Label"
                            Caption ="Special Instr Product2"
                            EventProcPrefix ="Special_Instr_Product2_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22920
                    Top =3360
                    Width =1920
                    Height =255
                    ColumnWidth =3000
                    TabIndex =79
                    Name ="Special Instr Product3"
                    ControlSource ="Special Instr Product3"
                    EventProcPrefix ="Special_Instr_Product3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =22140
                            Top =3360
                            Width =720
                            Height =255
                            Name ="Special Instr Product3_Label"
                            Caption ="Special Instr Product3"
                            EventProcPrefix ="Special_Instr_Product3_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25680
                    Top =120
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =80
                    Name ="Special Instr Product4"
                    ControlSource ="Special Instr Product4"
                    EventProcPrefix ="Special_Instr_Product4"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =120
                            Width =720
                            Height =255
                            Name ="Special Instr Product4_Label"
                            Caption ="Special Instr Product4"
                            EventProcPrefix ="Special_Instr_Product4_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25680
                    Top =660
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =81
                    Name ="Label Type"
                    ControlSource ="Label Type"
                    EventProcPrefix ="Label_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =660
                            Width =720
                            Height =255
                            Name ="Label Type_Label"
                            Caption ="Label Type"
                            EventProcPrefix ="Label_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25680
                    Top =1020
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =82
                    Name ="\"E\" Number"
                    ControlSource ="\"E\" Number"
                    EventProcPrefix ="Ctl_E__Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =1020
                            Width =720
                            Height =255
                            Name ="\"E\" Number_Label"
                            Caption ="\"E\" Number"
                            EventProcPrefix ="Ctl_E__Number_Label"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =25680
                    Top =1380
                    TabIndex =83
                    Name ="load into oracle 11?"
                    ControlSource ="load into oracle 11?"
                    EventProcPrefix ="load_into_oracle_11_"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =1380
                            Width =720
                            Height =255
                            Name ="load into oracle 11?_Label"
                            Caption ="load into oracle 11?"
                            EventProcPrefix ="load_into_oracle_11__Label"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =25680
                    Top =1740
                    TabIndex =84
                    Name ="calculate bill?"
                    ControlSource ="calculate bill?"
                    EventProcPrefix ="calculate_bill_"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =1740
                            Width =720
                            Height =255
                            Name ="calculate bill?_Label"
                            Caption ="calculate bill?"
                            EventProcPrefix ="calculate_bill__Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25680
                    Top =2100
                    Width =1920
                    Height =255
                    ColumnWidth =1935
                    TabIndex =85
                    Name ="Standard Operation"
                    ControlSource ="Standard Operation"
                    EventProcPrefix ="Standard_Operation"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =2100
                            Width =720
                            Height =255
                            Name ="Standard Operation_Label"
                            Caption ="Standard Operation"
                            EventProcPrefix ="Standard_Operation_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25680
                    Top =2460
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =86
                    Name ="OracleStatus"
                    ControlSource ="OracleStatus"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =2460
                            Width =720
                            Height =255
                            Name ="OracleStatus_Label"
                            Caption ="OracleStatus"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25680
                    Top =2820
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =87
                    Name ="Aramid1Type"
                    ControlSource ="Aramid1Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =2820
                            Width =720
                            Height =255
                            Name ="Aramid1Type_Label"
                            Caption ="Aramid1Type"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25680
                    Top =3360
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =88
                    Name ="Aramid1Count"
                    ControlSource ="Aramid1Count"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24900
                            Top =3360
                            Width =720
                            Height =255
                            Name ="Aramid1Count_Label"
                            Caption ="Aramid1Count"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =28440
                    Top =120
                    Width =1920
                    Height =450
                    ColumnWidth =3000
                    TabIndex =89
                    Name ="Aramid2Type"
                    ControlSource ="Aramid2Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27660
                            Top =120
                            Width =720
                            Height =255
                            Name ="Aramid2Type_Label"
                            Caption ="Aramid2Type"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =28440
                    Top =660
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =90
                    Name ="Aramid2Count"
                    ControlSource ="Aramid2Count"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27660
                            Top =660
                            Width =720
                            Height =255
                            Name ="Aramid2Count_Label"
                            Caption ="Aramid2Count"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End

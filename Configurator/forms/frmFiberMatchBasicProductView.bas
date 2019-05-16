Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =30660
    DatasheetFontHeight =10
    ItemSuffix =150
    Left =300
    Top =1260
    Right =19050
    Bottom =10920
    DatasheetGridlinesColor =12632256
    Filter ="((([Basic Product Construction].Base=\"KQ006-6.4\"))) AND (([Basic Product Const"
        "ruction].[New Oracle Part #] Like \"PK02296*\"))"
    OrderBy ="[Basic Product Construction].[New Oracle Part #]"
    RecSrcDt = Begin
        0xb5d3e9632b99e440
    End
    RecordSource ="qryfrmFiberMatchExistingFind"
    Caption ="qryfrmFiberMatchExistingFind"
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
            Height =4170
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    Name ="New Oracle Part #"
                    ControlSource ="New Oracle Part #"
                    EventProcPrefix ="New_Oracle_Part__"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =780
                            Height =255
                            Name ="New Oracle Part #_Label"
                            Caption ="New Oracle Part #"
                            EventProcPrefix ="New_Oracle_Part___Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =660
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="Item No"
                    ControlSource ="Item No"
                    EventProcPrefix ="Item_No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =660
                            Width =780
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
                    Left =900
                    Top =1200
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Customer"
                    ControlSource ="Customer"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =780
                            Height =255
                            Name ="Customer_Label"
                            Caption ="Customer"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =1740
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="Customer Part#"
                    ControlSource ="Customer Part#"
                    EventProcPrefix ="Customer_Part_"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1740
                            Width =780
                            Height =255
                            Name ="Customer Part#_Label"
                            Caption ="Customer Part#"
                            EventProcPrefix ="Customer_Part__Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =2280
                    Width =465
                    Height =255
                    ColumnWidth =585
                    TabIndex =4
                    Name ="Active"
                    ControlSource ="Active"
                    Format ="Yes/No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2280
                            Width =780
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
                    Left =900
                    Top =2640
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="Reason"
                    ControlSource ="Reason"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2640
                            Width =780
                            Height =255
                            Name ="Reason_Label"
                            Caption ="Reason"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =3180
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =6
                    Name ="calculate bill?"
                    ControlSource ="calculate bill?"
                    EventProcPrefix ="calculate_bill_"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3180
                            Width =780
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
                    Left =900
                    Top =3540
                    Width =2160
                    Height =255
                    ColumnWidth =2310
                    TabIndex =7
                    Name ="Base"
                    ControlSource ="Base"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3540
                            Width =780
                            Height =255
                            Name ="Base_Label"
                            Caption ="Base"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3960
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="Cable Type"
                    ControlSource ="Cable Type"
                    EventProcPrefix ="Cable_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =120
                            Width =780
                            Height =255
                            Name ="Cable Type_Label"
                            Caption ="Cable Type"
                            EventProcPrefix ="Cable_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3960
                    Top =660
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =9
                    Name ="Revision Date"
                    ControlSource ="Revision Date"
                    EventProcPrefix ="Revision_Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =660
                            Width =780
                            Height =255
                            Name ="Revision Date_Label"
                            Caption ="Revision Date"
                            EventProcPrefix ="Revision_Date_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3960
                    Top =1020
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="Revision Letter"
                    ControlSource ="Revision Letter"
                    EventProcPrefix ="Revision_Letter"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =1020
                            Width =780
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
                    Left =3960
                    Top =1560
                    Width =2160
                    Height =840
                    ColumnWidth =3000
                    TabIndex =11
                    Name ="Revision History"
                    ControlSource ="Revision History"
                    EventProcPrefix ="Revision_History"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =1560
                            Width =780
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
                    Left =3960
                    Top =2520
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =12
                    Name ="Standard Operation"
                    ControlSource ="Standard Operation"
                    EventProcPrefix ="Standard_Operation"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =2520
                            Width =780
                            Height =255
                            Name ="Standard Operation_Label"
                            Caption ="Standard Operation"
                            EventProcPrefix ="Standard_Operation_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3960
                    Top =3060
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =13
                    Name ="Jacket Material"
                    ControlSource ="Jacket Material"
                    EventProcPrefix ="Jacket_Material"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =3060
                            Width =780
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
                    Left =3960
                    Top =3600
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =14
                    Name ="Jacket Color"
                    ControlSource ="Jacket Color"
                    EventProcPrefix ="Jacket_Color"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =3600
                            Width =780
                            Height =255
                            Name ="Jacket Color_Label"
                            Caption ="Jacket Color"
                            EventProcPrefix ="Jacket_Color_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =15
                    Name ="Color Chip ID"
                    ControlSource ="Color Chip ID"
                    EventProcPrefix ="Color_Chip_ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =120
                            Width =780
                            Height =255
                            Name ="Color Chip ID_Label"
                            Caption ="Color Chip ID"
                            EventProcPrefix ="Color_Chip_ID_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =660
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =16
                    Name ="Nominal OD"
                    ControlSource ="Nominal OD"
                    EventProcPrefix ="Nominal_OD"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =660
                            Width =780
                            Height =255
                            Name ="Nominal OD_Label"
                            Caption ="Nominal OD"
                            EventProcPrefix ="Nominal_OD_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =1200
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =17
                    Name ="OD Tol  (+)"
                    ControlSource ="OD Tol  (+)"
                    EventProcPrefix ="OD_Tol_____"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =1200
                            Width =780
                            Height =255
                            Name ="OD Tol  (+)_Label"
                            Caption ="OD Tol  (+)"
                            EventProcPrefix ="OD_Tol______Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =1740
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =18
                    Name ="OD Tol  (-)1"
                    ControlSource ="OD Tol  (-)"
                    EventProcPrefix ="OD_Tol_____1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =1740
                            Width =780
                            Height =255
                            Name ="OD Tol  (-)1_Label"
                            Caption ="OD Tol  (-)1"
                            EventProcPrefix ="OD_Tol_____1_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =2280
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =19
                    Name ="Height"
                    ControlSource ="Height"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =2280
                            Width =780
                            Height =255
                            Name ="Height_Label"
                            Caption ="Height"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =2820
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =20
                    Name ="Height Tol (+)"
                    ControlSource ="Height Tol (+)"
                    EventProcPrefix ="Height_Tol____"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =2820
                            Width =780
                            Height =255
                            Name ="Height Tol (+)_Label"
                            Caption ="Height Tol (+)"
                            EventProcPrefix ="Height_Tol_____Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =3360
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =21
                    Name ="Height Tol (-)2"
                    ControlSource ="Height Tol (-)"
                    EventProcPrefix ="Height_Tol____2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =3360
                            Width =780
                            Height =255
                            Name ="Height Tol (-)2_Label"
                            Caption ="Height Tol (-)2"
                            EventProcPrefix ="Height_Tol____2_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10080
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =22
                    Name ="Nominal Wall"
                    ControlSource ="Nominal Wall"
                    EventProcPrefix ="Nominal_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9240
                            Top =120
                            Width =780
                            Height =255
                            Name ="Nominal Wall_Label"
                            Caption ="Nominal Wall"
                            EventProcPrefix ="Nominal_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10080
                    Top =660
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =23
                    Name ="Max Ave Wall"
                    ControlSource ="Max Ave Wall"
                    EventProcPrefix ="Max_Ave_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9240
                            Top =660
                            Width =780
                            Height =255
                            Name ="Max Ave Wall_Label"
                            Caption ="Max Ave Wall"
                            EventProcPrefix ="Max_Ave_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10080
                    Top =1200
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =24
                    Name ="Min Ave Wall"
                    ControlSource ="Min Ave Wall"
                    EventProcPrefix ="Min_Ave_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9240
                            Top =1200
                            Width =780
                            Height =255
                            Name ="Min Ave Wall_Label"
                            Caption ="Min Ave Wall"
                            EventProcPrefix ="Min_Ave_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10080
                    Top =1740
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =25
                    Name ="Min Spot Wall"
                    ControlSource ="Min Spot Wall"
                    EventProcPrefix ="Min_Spot_Wall"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9240
                            Top =1740
                            Width =780
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
                    Left =10080
                    Top =2280
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =26
                    Name ="Print Reel No"
                    ControlSource ="Print Reel No"
                    Format ="Yes/No"
                    EventProcPrefix ="Print_Reel_No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9240
                            Top =2280
                            Width =780
                            Height =255
                            Name ="Print Reel No_Label"
                            Caption ="Print Reel No"
                            EventProcPrefix ="Print_Reel_No_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10080
                    Top =2640
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =27
                    Name ="Print Item No"
                    ControlSource ="Print Item No"
                    Format ="Yes/No"
                    EventProcPrefix ="Print_Item_No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9240
                            Top =2640
                            Width =780
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
                    Left =10080
                    Top =3000
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =28
                    Name ="Std Subunit Print"
                    ControlSource ="Std Subunit Print"
                    EventProcPrefix ="Std_Subunit_Print"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9240
                            Top =3000
                            Width =780
                            Height =255
                            Name ="Std Subunit Print_Label"
                            Caption ="Std Subunit Print"
                            EventProcPrefix ="Std_Subunit_Print_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10080
                    Top =3360
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =29
                    Name ="Print Type (base)"
                    ControlSource ="Print Type (base)"
                    EventProcPrefix ="Print_Type__base_"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9240
                            Top =3360
                            Width =780
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
                    Left =13140
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =30
                    Name ="Print Line 1"
                    ControlSource ="Print Line 1"
                    EventProcPrefix ="Print_Line_1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12300
                            Top =120
                            Width =780
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
                    Left =13140
                    Top =660
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =31
                    Name ="Print Line 2"
                    ControlSource ="Print Line 2"
                    EventProcPrefix ="Print_Line_2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12300
                            Top =660
                            Width =780
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
                    Left =13140
                    Top =1200
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =32
                    Name ="Print Line 3"
                    ControlSource ="Print Line 3"
                    EventProcPrefix ="Print_Line_3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12300
                            Top =1200
                            Width =780
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
                    Left =13140
                    Top =1740
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =33
                    Name ="Print Line 4"
                    ControlSource ="Print Line 4"
                    EventProcPrefix ="Print_Line_4"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12300
                            Top =1740
                            Width =780
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
                    Left =13140
                    Top =2280
                    Width =2160
                    Height =255
                    ColumnWidth =2310
                    TabIndex =34
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="NewPrintLine4"
                    ControlSource ="NewPrintLine4"
                    RowSourceType ="Table/Query"
                    RowSource ="tblLine4PrintCode"
                    ColumnWidths ="14;4032;1008"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12300
                            Top =2280
                            Width =780
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
                    Left =13140
                    Top =2640
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =35
                    Name ="Print Spacing"
                    ControlSource ="Print Spacing"
                    EventProcPrefix ="Print_Spacing"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12300
                            Top =2640
                            Width =780
                            Height =255
                            Name ="Print Spacing_Label"
                            Caption ="Print Spacing"
                            EventProcPrefix ="Print_Spacing_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13140
                    Top =3180
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =36
                    Name ="PrintNotes"
                    ControlSource ="PrintNotes"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12300
                            Top =3180
                            Width =780
                            Height =255
                            Name ="PrintNotes_Label"
                            Caption ="PrintNotes"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16200
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =37
                    Name ="Listing Company"
                    ControlSource ="Listing Company"
                    EventProcPrefix ="Listing_Company"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15360
                            Top =120
                            Width =780
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
                    Left =16200
                    Top =660
                    Width =2160
                    Height =255
                    ColumnWidth =2310
                    TabIndex =38
                    Name ="UL-ETL-Listing"
                    ControlSource ="UL-ETL-Listing"
                    EventProcPrefix ="UL_ETL_Listing"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15360
                            Top =660
                            Width =780
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
                    Left =16200
                    Top =1020
                    Width =2160
                    Height =255
                    ColumnWidth =2310
                    TabIndex =39
                    Name ="UL-ETL-Const"
                    ControlSource ="UL-ETL-Const"
                    EventProcPrefix ="UL_ETL_Const"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15360
                            Top =1020
                            Width =780
                            Height =255
                            Name ="UL-ETL-Const_Label"
                            Caption ="UL-ETL-Const"
                            EventProcPrefix ="UL_ETL_Const_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16200
                    Top =1380
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =40
                    Name ="UL-ETL-Section"
                    ControlSource ="UL-ETL-Section"
                    EventProcPrefix ="UL_ETL_Section"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15360
                            Top =1380
                            Width =780
                            Height =255
                            Name ="UL-ETL-Section_Label"
                            Caption ="UL-ETL-Section"
                            EventProcPrefix ="UL_ETL_Section_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16200
                    Top =1920
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =41
                    Name ="\"E\" Number"
                    ControlSource ="\"E\" Number"
                    EventProcPrefix ="Ctl_E__Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15360
                            Top =1920
                            Width =780
                            Height =255
                            Name ="\"E\" Number_Label"
                            Caption ="\"E\" Number"
                            EventProcPrefix ="Ctl_E__Number_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16200
                    Top =2460
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =42
                    Name ="Label Type"
                    ControlSource ="Label Type"
                    EventProcPrefix ="Label_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15360
                            Top =2460
                            Width =780
                            Height =255
                            Name ="Label Type_Label"
                            Caption ="Label Type"
                            EventProcPrefix ="Label_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16200
                    Top =3000
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =43
                    Name ="Aramid Type"
                    ControlSource ="Aramid Type"
                    EventProcPrefix ="Aramid_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15360
                            Top =3000
                            Width =780
                            Height =255
                            Name ="Aramid Type_Label"
                            Caption ="Aramid Type"
                            EventProcPrefix ="Aramid_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16200
                    Top =3540
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =44
                    Name ="Number of ends"
                    ControlSource ="Number of ends"
                    EventProcPrefix ="Number_of_ends"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15360
                            Top =3540
                            Width =780
                            Height =255
                            Name ="Number of ends_Label"
                            Caption ="Number of ends"
                            EventProcPrefix ="Number_of_ends_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19260
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =45
                    Name ="Aramid Type2"
                    ControlSource ="Aramid Type2"
                    EventProcPrefix ="Aramid_Type2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18420
                            Top =120
                            Width =780
                            Height =255
                            Name ="Aramid Type2_Label"
                            Caption ="Aramid Type2"
                            EventProcPrefix ="Aramid_Type2_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19260
                    Top =660
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =46
                    Name ="Number of ends2"
                    ControlSource ="Number of ends2"
                    EventProcPrefix ="Number_of_ends2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18420
                            Top =660
                            Width =780
                            Height =255
                            Name ="Number of ends2_Label"
                            Caption ="Number of ends2"
                            EventProcPrefix ="Number_of_ends2_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19260
                    Top =1200
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =47
                    Name ="Unit Chips Type"
                    ControlSource ="Unit Chips Type"
                    EventProcPrefix ="Unit_Chips_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18420
                            Top =1200
                            Width =780
                            Height =255
                            Name ="Unit Chips Type_Label"
                            Caption ="Unit Chips Type"
                            EventProcPrefix ="Unit_Chips_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19260
                    Top =1740
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =48
                    Name ="Unit Ripcord"
                    ControlSource ="Unit Ripcord"
                    EventProcPrefix ="Unit_Ripcord"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18420
                            Top =1740
                            Width =780
                            Height =255
                            Name ="Unit Ripcord_Label"
                            Caption ="Unit Ripcord"
                            EventProcPrefix ="Unit_Ripcord_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19260
                    Top =2280
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =49
                    Name ="Unit CM"
                    ControlSource ="Unit CM"
                    EventProcPrefix ="Unit_CM"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18420
                            Top =2280
                            Width =780
                            Height =255
                            Name ="Unit CM_Label"
                            Caption ="Unit CM"
                            EventProcPrefix ="Unit_CM_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19260
                    Top =2820
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =50
                    Name ="Lay Length"
                    ControlSource ="Lay Length"
                    EventProcPrefix ="Lay_Length"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18420
                            Top =2820
                            Width =780
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
                    Left =19260
                    Top =3360
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =51
                    Name ="Talc"
                    ControlSource ="Talc"
                    Format ="Yes/No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18420
                            Top =3360
                            Width =780
                            Height =255
                            Name ="Talc_Label"
                            Caption ="Talc"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19260
                    Top =3720
                    Width =465
                    Height =255
                    ColumnWidth =465
                    TabIndex =52
                    Name ="EZ Strip"
                    ControlSource ="EZ Strip"
                    Format ="Yes/No"
                    EventProcPrefix ="EZ_Strip"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18420
                            Top =3720
                            Width =780
                            Height =255
                            Name ="EZ Strip_Label"
                            Caption ="EZ Strip"
                            EventProcPrefix ="EZ_Strip_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22320
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =53
                    Name ="TB Material"
                    ControlSource ="TB Material"
                    EventProcPrefix ="TB_Material"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21480
                            Top =120
                            Width =780
                            Height =255
                            Name ="TB Material_Label"
                            Caption ="TB Material"
                            EventProcPrefix ="TB_Material_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22320
                    Top =660
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =54
                    Name ="TB Chips Type"
                    ControlSource ="TB Chips Type"
                    EventProcPrefix ="TB_Chips_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21480
                            Top =660
                            Width =780
                            Height =255
                            Name ="TB Chips Type_Label"
                            Caption ="TB Chips Type"
                            EventProcPrefix ="TB_Chips_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22320
                    Top =1200
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =55
                    Name ="TB Nominal OD"
                    ControlSource ="TB Nominal OD"
                    EventProcPrefix ="TB_Nominal_OD"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21480
                            Top =1200
                            Width =780
                            Height =255
                            Name ="TB Nominal OD_Label"
                            Caption ="TB Nominal OD"
                            EventProcPrefix ="TB_Nominal_OD_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22320
                    Top =1740
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =56
                    Name ="TB OD Tol  (+)"
                    ControlSource ="TB OD Tol  (+)"
                    EventProcPrefix ="TB_OD_Tol_____"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21480
                            Top =1740
                            Width =780
                            Height =255
                            Name ="TB OD Tol  (+)_Label"
                            Caption ="TB OD Tol  (+)"
                            EventProcPrefix ="TB_OD_Tol______Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22320
                    Top =2280
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =57
                    Name ="TB OD Tol  (-)3"
                    ControlSource ="TB OD Tol  (-)"
                    EventProcPrefix ="TB_OD_Tol_____3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21480
                            Top =2280
                            Width =780
                            Height =255
                            Name ="TB OD Tol  (-)3_Label"
                            Caption ="TB OD Tol  (-)3"
                            EventProcPrefix ="TB_OD_Tol_____3_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22320
                    Top =2820
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =58
                    Name ="TB Color Series"
                    ControlSource ="TB Color Series"
                    EventProcPrefix ="TB_Color_Series"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21480
                            Top =2820
                            Width =780
                            Height =255
                            Name ="TB Color Series_Label"
                            Caption ="TB Color Series"
                            EventProcPrefix ="TB_Color_Series_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =22320
                    Top =3360
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =59
                    Name ="Fiber Type"
                    ControlSource ="Fiber Type"
                    EventProcPrefix ="Fiber_Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21480
                            Top =3360
                            Width =780
                            Height =255
                            Name ="Fiber Type_Label"
                            Caption ="Fiber Type"
                            EventProcPrefix ="Fiber_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25380
                    Top =120
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =60
                    Name ="Fiber Oracle item"
                    ControlSource ="Fiber Oracle item"
                    EventProcPrefix ="Fiber_Oracle_item"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24540
                            Top =120
                            Width =780
                            Height =255
                            Name ="Fiber Oracle item_Label"
                            Caption ="Fiber Oracle item"
                            EventProcPrefix ="Fiber_Oracle_item_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25380
                    Top =660
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =61
                    Name ="1st Req Freq"
                    ControlSource ="1st Req Freq"
                    EventProcPrefix ="Ctl1st_Req_Freq"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24540
                            Top =660
                            Width =780
                            Height =255
                            Name ="1st Req Freq_Label"
                            Caption ="1st Req Freq"
                            EventProcPrefix ="Ctl1st_Req_Freq_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25380
                    Top =1200
                    Width =2160
                    Height =255
                    ColumnWidth =2310
                    TabIndex =62
                    Name ="1st Max Atten"
                    ControlSource ="1st Max Atten"
                    EventProcPrefix ="Ctl1st_Max_Atten"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24540
                            Top =1200
                            Width =780
                            Height =255
                            Name ="1st Max Atten_Label"
                            Caption ="1st Max Atten"
                            EventProcPrefix ="Ctl1st_Max_Atten_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25380
                    Top =1560
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =63
                    Name ="1st Min BandW"
                    ControlSource ="1st Min BandW"
                    EventProcPrefix ="Ctl1st_Min_BandW"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24540
                            Top =1560
                            Width =780
                            Height =255
                            Name ="1st Min BandW_Label"
                            Caption ="1st Min BandW"
                            EventProcPrefix ="Ctl1st_Min_BandW_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25380
                    Top =2100
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =64
                    Name ="2nd Req Freq"
                    ControlSource ="2nd Req Freq"
                    EventProcPrefix ="Ctl2nd_Req_Freq"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24540
                            Top =2100
                            Width =780
                            Height =255
                            Name ="2nd Req Freq_Label"
                            Caption ="2nd Req Freq"
                            EventProcPrefix ="Ctl2nd_Req_Freq_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25380
                    Top =2640
                    Width =2160
                    Height =255
                    ColumnWidth =2310
                    TabIndex =65
                    Name ="2nd Max Atten"
                    ControlSource ="2nd Max Atten"
                    EventProcPrefix ="Ctl2nd_Max_Atten"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24540
                            Top =2640
                            Width =780
                            Height =255
                            Name ="2nd Max Atten_Label"
                            Caption ="2nd Max Atten"
                            EventProcPrefix ="Ctl2nd_Max_Atten_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25380
                    Top =3000
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =66
                    Name ="2nd Min BandW"
                    ControlSource ="2nd Min BandW"
                    EventProcPrefix ="Ctl2nd_Min_BandW"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24540
                            Top =3000
                            Width =780
                            Height =255
                            Name ="2nd Min BandW_Label"
                            Caption ="2nd Min BandW"
                            EventProcPrefix ="Ctl2nd_Min_BandW_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25380
                    Top =3540
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =67
                    Name ="SM 1300 Max Atten"
                    ControlSource ="SM 1300 Max Atten"
                    EventProcPrefix ="SM_1300_Max_Atten"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =24540
                            Top =3540
                            Width =780
                            Height =255
                            Name ="SM 1300 Max Atten_Label"
                            Caption ="SM 1300 Max Atten"
                            EventProcPrefix ="SM_1300_Max_Atten_Label"
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
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =68
                    Name ="SM 1550 Max Atten"
                    ControlSource ="SM 1550 Max Atten"
                    EventProcPrefix ="SM_1550_Max_Atten"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27600
                            Top =120
                            Width =780
                            Height =255
                            Name ="SM 1550 Max Atten_Label"
                            Caption ="SM 1550 Max Atten"
                            EventProcPrefix ="SM_1550_Max_Atten_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =28440
                    Top =660
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =69
                    Name ="Special Instr Product1"
                    ControlSource ="Special Instr Product1"
                    EventProcPrefix ="Special_Instr_Product1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27600
                            Top =660
                            Width =780
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
                    Left =28440
                    Top =1200
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =70
                    Name ="zzSpecial Instr Product3"
                    ControlSource ="zzSpecial Instr Product3"
                    EventProcPrefix ="zzSpecial_Instr_Product3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27600
                            Top =1200
                            Width =780
                            Height =255
                            Name ="zzSpecial Instr Product3_Label"
                            Caption ="zzSpecial Instr Product3"
                            EventProcPrefix ="zzSpecial_Instr_Product3_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =28440
                    Top =1740
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =71
                    Name ="zzSpecial Instr Product4"
                    ControlSource ="zzSpecial Instr Product4"
                    EventProcPrefix ="zzSpecial_Instr_Product4"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27600
                            Top =1740
                            Width =780
                            Height =255
                            Name ="zzSpecial Instr Product4_Label"
                            Caption ="zzSpecial Instr Product4"
                            EventProcPrefix ="zzSpecial_Instr_Product4_Label"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =28440
                    Top =2280
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =72
                    Name ="zzOracle Part#"
                    ControlSource ="zzOracle Part#"
                    EventProcPrefix ="zzOracle_Part_"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27600
                            Top =2280
                            Width =780
                            Height =255
                            Name ="zzOracle Part#_Label"
                            Caption ="zzOracle Part#"
                            EventProcPrefix ="zzOracle_Part__Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =28440
                    Top =2820
                    Width =2160
                    Height =255
                    ColumnWidth =2310
                    TabIndex =73
                    Name ="OracleStatus"
                    ControlSource ="OracleStatus"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27600
                            Top =2820
                            Width =780
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
                    Left =28440
                    Top =3180
                    Width =2160
                    Height =450
                    ColumnWidth =3000
                    TabIndex =74
                    Name ="PSS Document #"
                    ControlSource ="PSS Document #"
                    EventProcPrefix ="PSS_Document__"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =27600
                            Top =3180
                            Width =780
                            Height =255
                            Name ="PSS Document #_Label"
                            Caption ="PSS Document #"
                            EventProcPrefix ="PSS_Document___Label"
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

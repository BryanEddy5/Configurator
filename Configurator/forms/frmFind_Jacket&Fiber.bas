Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =31590
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =270
    Top =600
    Right =19050
    Bottom =4095
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xad407d329687e440
    End
    RecordSource ="qry_Find_Jacket&Fiber"
    Caption ="frmFind_Jacket&Fiber"
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =1920
                    Height =240
                    Name ="Item_No_Label"
                    Caption ="Item_No"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2040
                    Top =60
                    Width =1920
                    Height =240
                    Name ="Base_Label"
                    Caption ="Base"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4020
                    Top =60
                    Width =3720
                    Height =240
                    Name ="Jacket Material_Label"
                    Caption ="Jacket Material"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Jacket_Material_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7800
                    Top =60
                    Width =3720
                    Height =240
                    Name ="Color Chip ID_Label"
                    Caption ="Color Chip ID"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Color_Chip_ID_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =11580
                    Top =60
                    Width =1920
                    Height =240
                    Name ="Fiber_Label"
                    Caption ="Fiber"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =13560
                    Top =60
                    Width =1920
                    Height =240
                    Name ="SM/MM_Label"
                    Caption ="SM/MM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SM_MM_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =15540
                    Top =60
                    Width =1920
                    Height =240
                    Name ="OM#_Label"
                    Caption ="OM#"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="OM__Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =17520
                    Top =60
                    Width =1920
                    Height =240
                    Name ="BIF_Label"
                    Caption ="BIF"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =19500
                    Top =60
                    Width =3720
                    Height =240
                    Name ="Description_Oracle_Label"
                    Caption ="Description_Oracle"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =23280
                    Top =60
                    Width =1920
                    Height =240
                    Name ="Jacket_Color_letter_Label"
                    Caption ="Jacket_Color_letter"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =25260
                    Top =60
                    Width =1920
                    Height =240
                    Name ="Jacket_Material_letter_Label"
                    Caption ="Jacket_Material_letter"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =27240
                    Top =60
                    Width =1920
                    Height =240
                    Name ="TBF_Letter_Label"
                    Caption ="TBF_Letter"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =29220
                    Top =60
                    Width =2310
                    Height =240
                    Name ="Fiber_Letter_Label"
                    Caption ="Fiber_Letter"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =570
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    Name ="Item_No"
                    ControlSource ="Item_No"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Base"
                    ControlSource ="Base"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4020
                    Top =60
                    Width =3720
                    Height =450
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Jacket Material"
                    ControlSource ="Jacket Material"
                    EventProcPrefix ="Jacket_Material"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7800
                    Top =60
                    Width =3720
                    Height =450
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="Color Chip ID"
                    ControlSource ="Color Chip ID"
                    EventProcPrefix ="Color_Chip_ID"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11580
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =4
                    Name ="Fiber"
                    ControlSource ="Fiber"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13560
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="SM/MM"
                    ControlSource ="SM/MM"
                    EventProcPrefix ="SM_MM"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15540
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="OM#"
                    ControlSource ="OM#"
                    EventProcPrefix ="OM_"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17520
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =7
                    Name ="BIF"
                    ControlSource ="BIF"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19500
                    Top =60
                    Width =3720
                    Height =450
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="Description_Oracle"
                    ControlSource ="Description_Oracle"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =23280
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =9
                    Name ="Jacket_Color_letter"
                    ControlSource ="Jacket_Color_letter"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =25260
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =10
                    Name ="Jacket_Material_letter"
                    ControlSource ="Jacket_Material_letter"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =27240
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    TabIndex =11
                    Name ="TBF_Letter"
                    ControlSource ="TBF_Letter"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =29220
                    Top =60
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =12
                    Name ="Fiber_Letter"
                    ControlSource ="Fiber_Letter"

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

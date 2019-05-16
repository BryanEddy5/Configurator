Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13740
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =1800
    Top =1125
    DatasheetGridlinesColor =12632256
    Filter ="([tbl_Fibers].[Fiber]=\"FBR00423\")"
    OrderBy ="[tbl_Fibers].[Fiber], [tbl_Fibers].[ITU], [tbl_Fibers].[Cost]"
    RecSrcDt = Begin
        0x992bf95c53c5e440
    End
    RecordSource ="qryFibersActive"
    Caption ="qryFibersActive"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000ac3500008601000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="FiberValue"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =930
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =3555
                    Height =510
                    FontSize =20
                    Name ="Label16"
                    Caption ="Configurator Fibers"
                End
            End
        End
        Begin PageHeader
            Height =420
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =1380
                    Height =300
                    Name ="FiberValue_Label"
                    Caption ="FiberValue"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1500
                    Top =60
                    Width =1380
                    Height =300
                    Name ="Fiber_Label"
                    Caption ="Fiber"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2940
                    Top =60
                    Width =3900
                    Height =300
                    Name ="Description_Oracle_Label"
                    Caption ="Description_Oracle"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6960
                    Top =60
                    Width =900
                    Height =300
                    Name ="ITU_Label"
                    Caption ="ITU"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7980
                    Top =60
                    Width =1380
                    Height =300
                    Name ="OracleFiberFamily_Label"
                    Caption ="OracleFiberFamily"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =9420
                    Top =60
                    Width =1380
                    Height =300
                    Name ="Vendor_Label"
                    Caption ="Vendor"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =10860
                    Top =60
                    Width =480
                    Height =300
                    Name ="BIF_Label"
                    Caption ="BIF"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =11520
                    Top =60
                    Width =1440
                    Height =300
                    Name ="SpecialNotes_Label"
                    Caption ="SpecialNotes"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =120
                    Top =360
                    Width =12840
                    Name ="Line19"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1380
                    Height =270
                    ColumnWidth =5040
                    Name ="FiberValue"
                    ControlSource ="FiberValue"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1500
                    Top =60
                    Width =1380
                    Height =270
                    ColumnWidth =2415
                    TabIndex =1
                    Name ="Fiber"
                    ControlSource ="Fiber"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2940
                    Top =60
                    Width =3900
                    Height =270
                    ColumnWidth =9540
                    TabIndex =2
                    Name ="Description_Oracle"
                    ControlSource ="Description_Oracle"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6960
                    Top =60
                    Width =900
                    Height =270
                    TabIndex =3
                    Name ="ITU"
                    ControlSource ="ITU"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7980
                    Top =60
                    Width =1380
                    Height =270
                    ColumnWidth =1770
                    TabIndex =4
                    Name ="OracleFiberFamily"
                    ControlSource ="OracleFiberFamily"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9420
                    Top =60
                    Width =1380
                    Height =270
                    TabIndex =5
                    Name ="Vendor"
                    ControlSource ="Vendor"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10860
                    Top =60
                    Width =420
                    Height =270
                    TabIndex =6
                    Name ="BIF"
                    ControlSource ="BIF"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =11520
                    Top =60
                    Height =270
                    ColumnWidth =1680
                    TabIndex =7
                    Name ="SpecialNotes"
                    ControlSource ="SpecialNotes"

                End
            End
        End
        Begin PageFooter
            Height =540
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text17"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7860
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text18"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin Line
                    BorderWidth =3
                    Left =60
                    Top =240
                    Width =12840
                    BorderColor =12632256
                    Name ="Line20"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End

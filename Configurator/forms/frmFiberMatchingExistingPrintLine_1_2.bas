Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1620
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =9255
    Top =3570
    Right =18240
    Bottom =8505
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x11f01bc8d396e440
    End
    RecordSource ="qryFiberMatchExistingPrintLine_1_2"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Visible = NotDefault
            Height =1140
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =420
                    Top =180
                    Width =720
                    Height =360
                    Name ="Text0"
                    ControlSource ="PrintLine1"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =420
                    Top =660
                    Width =720
                    Height =360
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="PrintLine2"

                End
            End
        End
    End
End

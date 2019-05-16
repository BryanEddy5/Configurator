Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =22
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =660
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =4365
    Top =2655
    Right =5970
    Bottom =2655
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa1bccafce08ce440
    End
    RecordSource ="qrySubUnitFindNextOracleNo"
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
            Height =900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =540
                    Height =360
                    Name ="ItemNo"
                    ControlSource ="Item No"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =420
                    Width =300
                    Height =360
                    TabIndex =1
                    Name ="NewOracle"
                    ControlSource ="NewOracle"

                End
            End
        End
    End
End

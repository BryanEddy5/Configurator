Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =7
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3720
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =5565
    Top =2175
    Right =7560
    Bottom =3060
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1df655f43394e440
    End
    RecordSource ="qryMatchSpeciaPrint"
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
            Height =1620
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =180
                    Width =3120
                    Name ="Text0"
                    ControlSource ="Special_1"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =540
                    Width =3120
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="Special_2"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =900
                    Width =3120
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="Special_3"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =1260
                    Width =3120
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="Special_4"

                End
            End
        End
    End
End

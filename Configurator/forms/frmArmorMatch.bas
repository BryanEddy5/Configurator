Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =7
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2100
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =8925
    Top =2130
    Right =11280
    Bottom =5775
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xca5f5b942c03e540
    End
    RecordSource ="qryArmorMatch"
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
            Height =3660
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =120
                    Height =360
                    Name ="MaxCore"
                    ControlSource ="MaxCoreOD"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =600
                    Height =360
                    TabIndex =1
                    Name ="ArmorNo"
                    ControlSource ="ArmorNo"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =1080
                    Height =360
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="JacketMaterial"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =1560
                    Height =360
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="NomJacketOD"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =2040
                    Height =480
                    TabIndex =4
                    Name ="OracleBase"
                    ControlSource ="OracleBase"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =2580
                    Height =480
                    TabIndex =5
                    Name ="NomOD"
                    ControlSource ="BaseOD"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =3180
                    Height =480
                    TabIndex =6
                    Name ="MaxJacketOD"
                    ControlSource ="MaxJacketOD"

                End
            End
        End
    End
End

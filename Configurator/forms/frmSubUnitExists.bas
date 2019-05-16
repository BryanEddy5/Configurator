Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =22
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1680
    DatasheetFontHeight =10
    ItemSuffix =3
    Right =1485
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x19159a983092e440
    End
    RecordSource ="qrytblSubunitExists"
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
            Height =660
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =180
                    Top =180
                    Width =840
                    Name ="Exists"
                    ControlSource ="Exists"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =180
                            Width =555
                            Height =240
                            Name ="Label2"
                            Caption ="Text1:"
                        End
                    End
                End
            End
        End
    End
End

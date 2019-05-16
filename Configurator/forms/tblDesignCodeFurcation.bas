Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4539
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =2850
    Top =2340
    Right =7680
    Bottom =3105
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2c129df94e9be440
    End
    RecordSource ="tblDesignCodeFurcationTube"
    Caption ="tblDesignCodeTB"
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
            Height =840
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =120
                    Width =2475
                    Name ="FurcationLetter"
                    ControlSource ="FurcationLetter"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =1185
                            Height =240
                            Name ="Label1"
                            Caption ="FurcationLetter"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =480
                    Width =2475
                    ColumnWidth =3540
                    TabIndex =1
                    Name ="FurcationTube"
                    ControlSource ="FurcationTube"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =1110
                            Height =240
                            Name ="Label3"
                            Caption ="FurcationTube"
                        End
                    End
                End
            End
        End
    End
End

Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4620
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =1785
    Top =2190
    Right =6420
    Bottom =3225
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x892206064e8ce440
    End
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin Section
            Height =720
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionButton
                    OverlapFlags =85
                    Left =180
                    Top =270
                    Name ="BasicCable"
                    AfterUpdate ="mcrLocationBasicCable"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =405
                            Top =240
                            Width =1575
                            Height =240
                            Name ="Label1"
                            Caption ="Basic (Cable/Sheath)"
                        End
                    End
                End
                Begin OptionButton
                    OverlapFlags =85
                    Left =2220
                    Top =270
                    TabIndex =1
                    Name ="BasicProduct"
                    AfterUpdate ="mcrLocationBasicProduct"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2445
                            Top =240
                            Width =1995
                            Height =240
                            Name ="Label3"
                            Caption ="Basic Product Construction"
                        End
                    End
                End
            End
        End
    End
End

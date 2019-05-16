Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    TabularFamily =220
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =1
    Left =11460
    Top =3405
    Right =19050
    Bottom =7260
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0cc85e9713abe440
    End
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
            Height =2880
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    Left =120
                    Top =1335
                    Width =1740
                    Height =255
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Nom Wall"
                    ControlSource ="NominalWall"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Nom_Wall"
                    AsianLineBreak =0

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =120
                            Top =1080
                            Width =1740
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Nom Wall Label"
                            Caption ="Nom Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Nom_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    Left =1860
                    Top =1335
                    Width =1740
                    Height =255
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Min Ave Wall"
                    ControlSource ="minavewall"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Min_Ave_Wall"
                    AsianLineBreak =0

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =1860
                            Top =1080
                            Width =1740
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Min Ave Wall Label"
                            Caption ="Min Ave Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Min_Ave_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    Left =3600
                    Top =1335
                    Width =1740
                    Height =255
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Min Spot Wall"
                    ControlSource ="Minspotwall"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Min_Spot_Wall"
                    AsianLineBreak =0

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =3600
                            Top =1080
                            Width =1740
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Min Spot Wall Label"
                            Caption ="Min Spot Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Min_Spot_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =5340
                    Top =1335
                    Width =1740
                    Height =255
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Max Ave Wall"
                    ControlSource ="Maxavewall"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Max_Ave_Wall"
                    AsianLineBreak =0

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =5340
                            Top =1080
                            Width =1740
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Max Ave Wall Label"
                            Caption ="Max Ave Wall"
                            FontName ="MS Sans Serif"
                            EventProcPrefix ="Max_Ave_Wall_Label"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =180
                    Width =2535
                    Height =600
                    FontSize =18
                    Name ="Label0"
                    Caption ="Mils Conversion"
                End
            End
        End
    End
End

Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    TabularFamily =48
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2460
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =4395
    Top =2910
    Right =8820
    Bottom =8505
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4eb09cabebc4e440
    End
    RecordSource ="qryDesignCodeODJacketMatch"
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
            Height =2460
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =480
                    Top =300
                    Width =840
                    Height =300
                    Name ="Location"
                    ControlSource ="CablePasses"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =480
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =1
                    Name ="Base"
                    ControlSource ="Base"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =480
                    Top =1140
                    Width =840
                    Height =300
                    TabIndex =2
                    Name ="Jacket"
                    ControlSource ="JacketMaterial"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =480
                    Top =1560
                    Width =840
                    Height =300
                    TabIndex =3
                    Name ="SubOD"
                    ControlSource ="SubOD"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =300
                    Width =840
                    Height =300
                    TabIndex =4
                    Name ="TB"
                    ControlSource ="TBType"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =5
                    Name ="OD"
                    ControlSource ="NominalOD"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =480
                    Top =1980
                    Width =840
                    Height =300
                    TabIndex =6
                    Name ="SubDesc"
                    ControlSource ="DesignCode"
                    BeforeUpdate ="SubDesc"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1140
                    Width =840
                    Height =300
                    TabIndex =7
                    Name ="Text15"
                    ControlSource ="DesignExtension"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1140
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1440
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1560
                    Width =840
                    Height =300
                    TabIndex =8
                    Name ="Text16"
                    ControlSource ="DesignTypeID"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1560
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1860
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1980
                    Width =840
                    Height =300
                    TabIndex =9
                    Name ="Text17"
                    ControlSource ="CableLevel1"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1980
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =2280
                End
            End
        End
    End
End

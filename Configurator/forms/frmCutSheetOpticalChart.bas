Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13980
    DatasheetFontHeight =11
    Left =6390
    Top =1995
    Right =19050
    Bottom =10920
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x2d803d7ef6c4e440
    End
    RecordSource ="SELECT tblCableTightBufferReference.TimeStamp FROM tblCableConstructionReference"
        "s INNER JOIN tblCableTightBufferReference ON tblCableConstructionReferences.TBTy"
        "pe = tblCableTightBufferReference.TBType;"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =-2147483643
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            ForeThemeColorIndex =2
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =360
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Top =60
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Oracle"
                    FontName ="Tahoma"
                    LayoutCachedTop =60
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =11400
                    Top =60
                    Width =2520
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Customer"
                    FontName ="Tahoma"
                    LayoutCachedLeft =11400
                    LayoutCachedTop =60
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1920
                    Top =60
                    Width =3720
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Catalog Number"
                    FontName ="Tahoma"
                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =7200
                    Top =60
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label24"
                    Caption ="Active"
                    FontName ="Tahoma"
                    LayoutCachedLeft =7200
                    LayoutCachedTop =60
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =7980
                    Top =60
                    Width =2760
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="Reason"
                    FontName ="Tahoma"
                    LayoutCachedLeft =7980
                    LayoutCachedTop =60
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =5760
                    Top =60
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Rev Date"
                    FontName ="Tahoma"
                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6600
                    Width =480
                    FontSize =8
                    ForeColor =-2147483630
                    Name ="btnSortDescending"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddd444d444dddd0dddd74ddd47ddd707dd ,
                        0xdd44444dddd000dddd74d47ddd70007dddd444dddd00000dddd747dddddd0ddd ,
                        0xdddd4ddddddd0ddddddddddddddd0dddd111111ddddd0dddd11ddd1ddddd0ddd ,
                        0xdd11dddddddd0dddddd11ddddddd0ddddddd11dddddd0dddd1ddd11ddddd0ddd ,
                        0xd111111ddddd0ddd
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =6600
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =16777215
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =426
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="Text6"
                    ControlSource ="Oracle"
                    FontName ="Tahoma"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11460
                    Top =60
                    Width =2520
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="Text18"
                    ControlSource ="Customer"
                    FontName ="Tahoma"

                    LayoutCachedLeft =11460
                    LayoutCachedTop =60
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1980
                    Top =60
                    Width =3720
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="Text22"
                    ControlSource ="Item No"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7260
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="Text25"
                    ControlSource ="Active"
                    Format ="Yes/No"
                    FontName ="Tahoma"

                    LayoutCachedLeft =7260
                    LayoutCachedTop =60
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10860
                    Width =426
                    Height =426
                    FontSize =8
                    TabIndex =4
                    ForeColor =-2147483630
                    Name ="btnFind"
                    Caption ="Command13"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadad000dadada ,
                        0xadadad0bb00dadaddadada0bb00adadaadadada00dadadaddadada0b00dadada ,
                        0xadadad0b00adadaddadada0b00dadadaadadad0bb00dadaddada00d0bb00dada ,
                        0xada0b00d0bb00daddad0b00ad0b00adaada0bb000bb00daddada0bbbbb00dada ,
                        0xadada000000dadad
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =10860
                    LayoutCachedWidth =11286
                    LayoutCachedHeight =426
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =16777215
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8040
                    Top =60
                    Width =2760
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="Text51"
                    ControlSource ="Reason"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5820
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text53"
                    ControlSource ="RevDate"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =60
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
        Begin FormFooter
            Height =360
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End

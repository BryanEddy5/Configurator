Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2794
    DatasheetFontHeight =11
    ItemSuffix =105
    Left =1275
    Top =2370
    Right =19050
    Bottom =10920
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xf0d29bde95c7e440
    End
    RecordSource ="SELECT tblBurnListing.Type, tblBurnListing.Base FROM tblBurnListingConstruction "
        "INNER JOIN tblBurnListing ON (tblBurnListingConstruction.ListingConstruction = t"
        "blBurnListing.ListingConstruction) AND (tblBurnListingConstruction.ListingSectio"
        "n = tblBurnListing.ListingSection) AND (tblBurnListingConstruction.ListingDocume"
        "nt = tblBurnListing.ListingTab) AND (tblBurnListingConstruction.ListingCompany ="
        " tblBurnListing.ListingCompany);"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =420
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =30
                    Width =2021
                    Height =330
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =15916467
                    ForeColor =3547148
                    Name ="Color"
                    ControlSource ="Type"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =2081
                    LayoutCachedHeight =360
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2145
                    Top =30
                    Width =611
                    Height =330
                    TabIndex =1
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =15916467
                    ForeColor =3547148
                    Name ="PrintLine3"
                    ControlSource ="Base"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =2145
                    LayoutCachedTop =30
                    LayoutCachedWidth =2756
                    LayoutCachedHeight =360
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
    End
End

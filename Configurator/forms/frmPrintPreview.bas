Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularFamily =220
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =21180
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =5235
    Top =4440
    Right =26415
    Bottom =5685
    DatasheetGridlinesColor =12632256
    Filter ="Oracle ='PS05528-01'"
    RecSrcDt = Begin
        0x852e44382bdfe440
    End
    RecordSource ="qryBasePrint"
    Caption ="Print Preview"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000001000000bc5200008403000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            BackColor =15915201
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =960
                    Top =180
                    Width =19920
                    Height =600
                    FontSize =11
                    Name ="Text5"
                    ControlSource ="=[PrintLine1] & \" \" & [PrintLine2] & IIf([Print Item No]<>0,\" \" & [Item No],"
                        "\" \") & \" \" & [PrintLine3] & \" \" & IIf([Print Reel No]<>0,\" Lot ##### \",\""
                        " \") & [PrintLine4]"

                    LayoutCachedLeft =960
                    LayoutCachedTop =180
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =780
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =-2147483633
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =840
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =840
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

Private Sub btnClose_Click()
    DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Me.Move 0, 0
    Me.btnClose.SetFocus

End Sub

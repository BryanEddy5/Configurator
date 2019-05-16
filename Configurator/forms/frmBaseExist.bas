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
    TabularFamily =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1920
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =4365
    Top =12480
    Right =6795
    Bottom =14580
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x39ef75566ed1e440
    End
    RecordSource ="qryItemNoSearchMatchBase"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =960
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =240
                    Top =240
                    Width =1560
                    Name ="BaseExist"
                    ControlSource ="DesignCode"

                    LayoutCachedLeft =240
                    LayoutCachedTop =240
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =240
                    Top =600
                    Width =1560
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="TBType"

                    LayoutCachedLeft =240
                    LayoutCachedTop =600
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =840
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


Private Sub Form_Open(Cancel As Integer)
move 50000, 50000
End Sub

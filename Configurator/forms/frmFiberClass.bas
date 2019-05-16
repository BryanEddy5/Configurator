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
    TabularFamily =88
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1080
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =5235
    Top =3780
    Right =8100
    Bottom =7275
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7e7f84ce75c0e440
    End
    RecordSource ="qryfrmFiberClass"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
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
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =900
                    Height =120
                    Name ="Text0"
                    ControlSource ="Fiber"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =180
                    Width =900
                    Height =180
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="OM#"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =420
                    Width =900
                    Height =120
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="SM/MM"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =600
                    Width =900
                    Height =180
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="BIF"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =840
                    Width =900
                    Height =180
                    TabIndex =4
                    Name ="ITU"
                    ControlSource ="ITU"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =1140
                    Width =900
                    Height =180
                    TabIndex =5
                    Name ="Text7"
                    ControlSource ="Vendor"

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
    Me.Visible = False
    Me.move 1000000, 1000000
End Sub

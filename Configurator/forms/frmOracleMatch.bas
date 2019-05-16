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
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1860
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =5235
    Top =885
    Right =8880
    Bottom =3480
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb9fa1379750fe540
    End
    RecordSource ="qryMainMenuSub"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Section
            Visible = NotDefault
            Height =540
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =180
                    Width =1560
                    Height =120
                    Name ="Oracle"
                    ControlSource ="[Oracle]"

                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =180
                    Top =120
                    Width =1560
                    Height =120
                    TabIndex =1
                    Name ="Exists"
                    ControlSource ="Exists"

                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =180
                    Top =180
                    Width =1560
                    Height =180
                    TabIndex =2
                    Name ="NewOracle"
                    ControlSource ="=Left([Oracle],9) & Mid([Oracle],10,1)+1 & Mid([Oracle],11,2)"

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
    Me.move 0, 100000
    Me.Visible = False
End Sub

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
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =22
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1380
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =5235
    Top =10635
    Right =8205
    Bottom =13035
    DatasheetGridlinesColor =12632256
    Filter ="(((([Standard Operation] Is Not Null))) AND (([Standard Operation] Is Not Null))"
        ") AND ((Not Location=1))"
    OrderBy ="[Standard Operation]"
    RecSrcDt = Begin
        0x8c1373b0aeb4e440
    End
    RecordSource ="qryfrmCreateItemNewOracleNo"
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
        Begin Section
            Visible = NotDefault
            Height =720
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Width =1080
                    Height =360
                    Name ="ItemNo"
                    ControlSource ="Oracle"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Top =360
                    Width =1080
                    Height =360
                    TabIndex =1
                    Name ="NewOracle"
                    ControlSource ="NewOracle"

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

Private Sub Form_Open(Cancel As Integer)
    Me.Visible = False
    Me.move 1000000, 1000000
    Me.Filter = ""
End Sub

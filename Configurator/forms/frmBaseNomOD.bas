Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =7
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2760
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =4365
    Top =1890
    Right =7065
    Bottom =2070
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x361fd766ef9ae440
    End
    RecordSource ="qryBaseNomODMatch"
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
            Height =1920
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =1140
                    Height =360
                    Name ="Text2"
                    ControlSource ="OD"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =60
                            Width =555
                            Height =240
                            Name ="Label3"
                            Caption ="Text2:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =300
                    Top =480
                    Width =1140
                    Height =360
                    TabIndex =1
                    Name ="Jacket"
                    ControlSource ="Jacket"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =480
                            Width =555
                            Height =240
                            Name ="Label5"
                            Caption ="Text2:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =300
                    Top =900
                    Width =1140
                    Height =360
                    TabIndex =2
                    Name ="Colo"
                    ControlSource ="Color"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =900
                            Width =555
                            Height =240
                            Name ="Label7"
                            Caption ="Text2:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =300
                    Top =1320
                    Width =1140
                    Height =360
                    TabIndex =3
                    Name ="ColorChip"
                    ControlSource ="ColorChip"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =1320
                            Width =555
                            Height =240
                            Name ="Label9"
                            Caption ="Text2:"
                        End
                    End
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
Me.move 100000, 100000

End Sub

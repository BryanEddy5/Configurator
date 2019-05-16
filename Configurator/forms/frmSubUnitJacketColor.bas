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
    TabularFamily =22
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2460
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =4365
    Top =4830
    Right =7575
    Bottom =7200
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x16e48fa1adc8e440
    End
    RecordSource ="qrySubUnitJacketColor"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
            Height =1680
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =720
                    Width =1320
                    Height =300
                    ColumnWidth =2010
                    Name ="JacketColor"
                    ControlSource ="Color"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =1140
                    Width =1320
                    Height =300
                    TabIndex =1
                    Name ="ColorChip"
                    ControlSource ="ColorChip"

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

Private Sub Form_Load()
Me.move 10000, 10000

End Sub

Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =29700
    DatasheetFontHeight =10
    ItemSuffix =1
    Right =19050
    Bottom =10920
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfbd84a5d1299e440
    End
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
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
        Begin Section
            Height =16380
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =600
                    Top =540
                    Width =8460
                    Height =2040
                    FontSize =24
                    ForeColor =8404992
                    Name ="Wait"
                    Caption ="Item is Loading, Please Wait..."
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
    DoCmd.Maximize
    DoEvents
End Sub

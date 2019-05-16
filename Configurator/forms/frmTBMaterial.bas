Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
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
    Width =1740
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =5235
    Top =3780
    Right =8760
    Bottom =7635
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb20c9832abc7e440
    End
    RecordSource ="qryfrmCreateItemTBMaterialMatch_ForTB"
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
            Height =1800
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =180
                    Width =1320
                    Height =300
                    Name ="TBMaterial"
                    ControlSource ="TB Material"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =600
                    Width =1320
                    Height =300
                    ColumnWidth =2010
                    TabIndex =1
                    Name ="TBChips"
                    ControlSource ="TB Chips Type"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =1020
                    Width =1320
                    Height =300
                    TabIndex =2
                    Name ="EZStrip"
                    ControlSource ="EZStrip"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =1440
                    Width =1320
                    Height =300
                    TabIndex =3
                    Name ="TBOD"
                    ControlSource ="TB Nominal OD"

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

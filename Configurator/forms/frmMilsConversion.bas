Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    TabularFamily =220
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7980
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =4365
    Top =5655
    Right =13020
    Bottom =10785
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0cc85e9713abe440
    End
    Caption ="Mils"
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
        Begin Line
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
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
            Height =2880
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    Left =480
                    Top =1335
                    Width =1740
                    Height =255
                    FontSize =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NominalWall"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            TextFontFamily =0
                            Left =480
                            Top =1080
                            Width =1740
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Nom Wall Label"
                            Caption ="Nom Wall"
                            FontName =""
                            EventProcPrefix ="Nom_Wall_Label"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    Left =3960
                    Top =1335
                    Width =1740
                    Height =255
                    FontSize =10
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MinAveWall"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3960
                    LayoutCachedTop =1335
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1590
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            TextFontFamily =0
                            Left =3960
                            Top =1080
                            Width =1740
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Min Ave Wall Label"
                            Caption ="Min Ave Wall"
                            FontName =""
                            EventProcPrefix ="Min_Ave_Wall_Label"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =1080
                            LayoutCachedWidth =5700
                            LayoutCachedHeight =1335
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =5700
                    Top =1335
                    Width =1740
                    Height =255
                    FontSize =10
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MinSpotWall"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =5700
                    LayoutCachedTop =1335
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =1590
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            TextFontFamily =0
                            Left =5700
                            Top =1080
                            Width =1740
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Min Spot Wall Label"
                            Caption ="Min Spot Wall"
                            FontName =""
                            EventProcPrefix ="Min_Spot_Wall_Label"
                            LayoutCachedLeft =5700
                            LayoutCachedTop =1080
                            LayoutCachedWidth =7440
                            LayoutCachedHeight =1335
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =2220
                    Top =1335
                    Width =1740
                    Height =255
                    FontSize =10
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="MaxAveWall"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2220
                    LayoutCachedTop =1335
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1590
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            TextFontFamily =0
                            Left =2220
                            Top =1080
                            Width =1740
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Max Ave Wall Label"
                            Caption ="Max Ave Wall"
                            FontName =""
                            EventProcPrefix ="Max_Ave_Wall_Label"
                            LayoutCachedLeft =2220
                            LayoutCachedTop =1080
                            LayoutCachedWidth =3960
                            LayoutCachedHeight =1335
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =540
                    Top =180
                    Width =2535
                    Height =600
                    FontSize =18
                    Name ="Label0"
                    Caption ="Mils Conversion"
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Top =1800
                    Width =7919
                    Name ="Line0"
                    LayoutCachedTop =1800
                    LayoutCachedWidth =7919
                    LayoutCachedHeight =1800
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5280
                    Top =2160
                    Width =840
                    Height =420
                    TabIndex =6
                    ForeColor =3547148
                    Name ="btnConvert"
                    Caption ="Convert"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5280
                    LayoutCachedTop =2160
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =2580
                    BorderColor =15916467
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2640
                    Top =2100
                    Width =480
                    Height =480
                    FontSize =24
                    FontWeight =500
                    Name ="Label6"
                    Caption ="="
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =2100
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =2580
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =3420
                    Top =2235
                    Width =1260
                    Height =255
                    FontSize =10
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txMils"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    OnKeyDown ="[Event Procedure]"
                    AsianLineBreak =0

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2235
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =2490
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            TextFontFamily =0
                            Left =3420
                            Top =1980
                            Width =1260
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label5"
                            Caption ="Mils"
                            FontName =""
                            LayoutCachedLeft =3420
                            LayoutCachedTop =1980
                            LayoutCachedWidth =4680
                            LayoutCachedHeight =2235
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =1020
                    Top =2235
                    Width =1260
                    Height =255
                    FontSize =10
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtMM"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    OnKeyDown ="[Event Procedure]"
                    AsianLineBreak =0

                    LayoutCachedLeft =1020
                    LayoutCachedTop =2235
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =2490
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            TextFontFamily =0
                            Left =1020
                            Top =1980
                            Width =1260
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label3"
                            Caption ="mm"
                            FontName =""
                            LayoutCachedLeft =1020
                            LayoutCachedTop =1980
                            LayoutCachedWidth =2280
                            LayoutCachedHeight =2235
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


Private Sub btnConvert_Click()

On Error GoTo ErrorHandler:

If IsNull(Me.txMils) = True Then
   Me.txMils = Round(Me.txtMM * 39.37, 3)
ElseIf IsNull(Me.txtMM) = True Then
    Me.txtMM = Round(Me.txMils / 39.37, 3)
Else
    Me.txMils = Round(Me.txtMM * 39.37, 3)
End If

ErrorHandler:
Exit Sub



End Sub


Private Sub txMils_AfterUpdate()

Me.txtMM = Null

End Sub

Private Sub txMils_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Me.btnConvert.SetFocus
        Call btnConvert_Click
    End If
End Sub

Private Sub txtMM_AfterUpdate()

Me.txMils = Null

End Sub

Private Sub txtMM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Me.btnConvert.SetFocus
        Call btnConvert_Click
    End If
End Sub

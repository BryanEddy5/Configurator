Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7080
    DatasheetFontHeight =10
    ItemSuffix =81
    Left =10575
    Top =2550
    Right =17910
    Bottom =9615
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xae875711f5e9e440
    End
    RecordSource ="SELECT tblBOMComponentExplode.BOMLevel, tblBOMComponentExplode.AssemblyItemNumbe"
        "r, tblBOMComponentExplode.Weight_kg_km FROM tblBOMComponentExplode ORDER BY tblB"
        "OMComponentExplode.BOMLevel, tblBOMComponentExplode.AssemblyItemNumber;"
    Caption ="Item Search"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyPress ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
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
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =2880
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =900
                    Top =2520
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Assembly"
                    LayoutCachedLeft =900
                    LayoutCachedTop =2520
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =2640
                    Top =2520
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="BOM Level"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =4320
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Weight (kg/km)"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2520
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =2820
                End
                Begin Line
                    OverlapFlags =85
                    Left =240
                    Top =1440
                    Width =6840
                    Name ="Line58"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1440
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =1440
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =1440
                    Top =1740
                    Width =3420
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItem"
                    Caption ="PE13968-00"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =1440
                    LayoutCachedTop =1740
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =2235
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =960
                    Top =300
                    Width =4740
                    Name ="Box35"
                    LayoutCachedLeft =960
                    LayoutCachedTop =300
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =1080
                    Top =420
                    Width =4500
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="BOM Component Explode"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =1080
                    LayoutCachedTop =420
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =915
                End
            End
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="txAssemblyItemNumber"
                    ControlSource ="AssemblyItemNumber"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =60
                    Width =1560
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txBomLevel"
                    ControlSource ="BOMLevel"

                    LayoutCachedLeft =2640
                    LayoutCachedTop =60
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =60
                    Width =975
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txWeight_kg_m"
                    ControlSource ="Weight_kg_km"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =60
                    LayoutCachedWidth =5295
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =960
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =300
                    Top =300
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =300
                    LayoutCachedTop =300
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =660
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
Option Compare Database
Option Explicit

Private Sub Base_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

Private Sub btnClose_Click()
DoCmd.Close acForm, Me.name
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

Private Sub txAssemblyItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

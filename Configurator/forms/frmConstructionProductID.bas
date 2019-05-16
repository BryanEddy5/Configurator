Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9600
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =10770
    Top =19965
    Right =19560
    Bottom =22155
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x93e5e9f2abfee440
    End
    RecordSource ="SELECT [Product Desrcriptions].* FROM [Product Desrcriptions]; "
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x0000000000000000000000000000000000000000ac2600007f08000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    Moveable =0
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =-2147483643
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            ForeThemeColorIndex =2
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =405
            BackColor =-2147483633
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =93
                    Width =5790
                    Height =405
                    FontSize =14
                    FontWeight =600
                    Name ="Label55"
                    Caption ="Product Primary Line / Sub Construction"
                    FontName ="MS Sans Serif"
                    LayoutCachedWidth =5790
                    LayoutCachedHeight =405
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5700
                    Width =720
                    FontSize =8
                    ForeColor =-2147483630
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5700
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =16777215
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6660
                    Width =720
                    FontSize =8
                    TabIndex =1
                    ForeColor =-2147483630
                    Name ="btnAdd"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =6660
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =16777215
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7620
                    Width =720
                    FontSize =8
                    TabIndex =2
                    ForeColor =-2147483630
                    Name ="btnSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =7620
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =16777215
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =120
                    Top =435
                    Width =1530
                    Height =300
                    FontSize =8
                    BackColor =-2147483643
                    Name ="ProductID"
                    ControlSource ="ProductID"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =435
                    LayoutCachedWidth =1650
                    LayoutCachedHeight =735
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =120
                            Width =1530
                            Height =315
                            FontSize =8
                            BorderColor =-2
                            Name ="ProductID_Label"
                            Caption ="ProductID"
                            FontName ="MS Sans Serif"
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =120
                            LayoutCachedWidth =1650
                            LayoutCachedHeight =435
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =1680
                    Top =435
                    Width =1200
                    Height =300
                    FontSize =8
                    TabIndex =1
                    BackColor =-2147483643
                    Name ="SubFiller"
                    ControlSource ="SubFiller"
                    Format ="True/False"
                    FontName ="MS Sans Serif"
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =435
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =735
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1680
                            Top =120
                            Width =1200
                            Height =315
                            FontSize =8
                            BorderColor =-2
                            Name ="SubFiller_Label"
                            Caption ="SubFiller"
                            FontName ="MS Sans Serif"
                            GridlineColor =10921638
                            LayoutCachedLeft =1680
                            LayoutCachedTop =120
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =435
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2880
                    Top =435
                    Width =2280
                    Height =300
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    Name ="txBase"
                    ControlSource ="Level 1 Product"
                    Format ="General Number"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =435
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =735
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =2880
                            Top =120
                            Width =2280
                            Height =315
                            FontSize =8
                            BorderColor =-2
                            Name ="Label4"
                            Caption ="Level 1 Product"
                            FontName ="MS Sans Serif"
                            GridlineColor =10921638
                            LayoutCachedLeft =2880
                            LayoutCachedTop =120
                            LayoutCachedWidth =5160
                            LayoutCachedHeight =435
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5340
                    Top =180
                    Width =426
                    Height =426
                    FontSize =8
                    TabIndex =3
                    ForeColor =-2147483630
                    Name ="btnConstruction"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000012000000120000000100180000000000f0030000110b0000110b0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffff6262620000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000636363ffffffffffffffff0000007070 ,
                        0x70bebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebe707070000000ffffffffffbebebe000000bebebeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffbebebe ,
                        0x000000ffffffffffbebebe000000bebebeffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffbebebe000000ffffffbeff ,
                        0xbebebe000000bebebeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffbebebe000000ffffff00ffbebebe000000bebe ,
                        0xbeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffbebebe000000ffffffffffbebebe000000bebebeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffbebebe ,
                        0x000000ffffff00ffbebebe000000bebebeffffffffffffffffffffffffffffff ,
                        0xffffff979797ffffffffffffffffffffffffffffffffffff0f0f0fffffffffff ,
                        0xbebebe000000bebebeffffffffffffffffffffffffffffff9696960000007171 ,
                        0x71ffffffffffffffffffffffffffffffffffffffffffffffbebebe000000bebe ,
                        0xbeffffffffffffffffffffffffffffffffffff707070000000717171ffffffff ,
                        0xffffffffffffffffffffffffffffffffbebebe000000bebebeffffffffffffff ,
                        0xffffffffffffffffffffffffffff707070000000717171ffffffffffffffffff ,
                        0x7f7f7fffffffffffbebebe000000bebebeffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff707070000000717171ffffffbebebe000000ffffffffff ,
                        0xbebebe000000bebebeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff707070000000717171bebebe000000ffffff3dffbebebe000000bebe ,
                        0xbeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff70 ,
                        0x70700000003d3d3d000000ffffff00ffffffff000000707070bebebebebebebe ,
                        0xbebebebebebebebeffffffffffffffffffffffffbebebebebebe3c3c3c000000 ,
                        0x000000ffffffbeffffffff626262000000000000000000000000000000000000 ,
                        0x101010ffffffffffff7f7f7f000000000000000000000000000000ffffffffff ,
                        0xffffffffffffffffffbebebebebebebebebebebebebebebeffffffffffffffff ,
                        0xffffffffbebebebebebebebebebebebeffffffffffffbeff
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="OpenIcon.bmp"

                    LayoutCachedLeft =5340
                    LayoutCachedTop =180
                    LayoutCachedWidth =5766
                    LayoutCachedHeight =606
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =16777215
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5940
                    Top =180
                    Width =426
                    Height =426
                    FontSize =8
                    TabIndex =4
                    ForeColor =-2147483630
                    Name ="btnDeleteRecord"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddddddd177ddddd77dd1ddd1177dddd17dd11dd7117ddd71ddd ,
                        0x111dd1177d117ddd1111d7117711dddd11111d11111ddddd1111dd71117ddddd ,
                        0x111d77111177dddd11d711dd71177ddd1dddddddd71177ddddddddddddd11ddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =5940
                    LayoutCachedTop =180
                    LayoutCachedWidth =6366
                    LayoutCachedHeight =606
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =16777215
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
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

Private Sub btnMils_Click()
On Error GoTo Error_Handler

DoCmd.OpenForm "frmMilsConversion"

Forms!frmMilsConversion!NominalWall = Round(CDbl(Me!frmCableConstruction.Form!NominalWall.NominalWall) * 39.37, 1)
Forms!frmMilsConversion!MinAveWall = Round(CDbl(Me!frmCableConstruction.Form!NominalWall.MinAveWall) * 39.37, 1)
Forms!frmMilsConversion!MinSpotWall = Round(CDbl(Me!frmCableConstruction.Form!NominalWall.MinSpotWall) * 39.37, 1)
Forms!frmMilsConversion!MaxAveWall = Round(CDbl(Me!frmCableConstruction.Form!NominalWall.MaxAveWall) * 39.37, 1)

Error_Exit:
Exit Sub

Error_Handler:
'MsgBox Err.Description & " " & Err.Number
Resume Error_Exit

End Sub

Private Sub btnAdd_Click()
Me.AllowAdditions = True
Call btnEdit_Click
'ProductID = Nz(Me.Parent!txSetupID)
End Sub

Private Sub btnConstruction_Click()

Call OpenAClient(3, Me.[Level 1 Product], 1)

End Sub


Private Sub btnDeleteRecord_Click()
Call DeleteARecord
End Sub

Private Sub btnEdit_Click()
Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If ctl.Locked <> False Then
            ctl.Locked = False
        End If
    End If
Next ctl
Me.ProductID.Locked = True
'ProductID = Nz(Me.Parent!txSetupID)
End Sub


Private Sub btnSave_Click()
Me.AllowAdditions = False
Me.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
Me.AllowAdditions = False
End Sub

Private Sub txBase_AfterUpdate()
ProductID = Nz(Me.Parent!txSetupID)
End Sub

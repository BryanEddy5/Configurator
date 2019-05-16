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
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18420
    DatasheetFontHeight =10
    ItemSuffix =69
    Left =4905
    Top =180
    Right =23325
    Bottom =14205
    DatasheetGridlinesColor =12632256
    Filter ="(frmFindProcessData.Location<>1 Or frmFindProcessData.Location Is Null)"
    OrderBy ="[Entry Date] DESC"
    RecSrcDt = Begin
        0x42878420d4cbe440
    End
    RecordSource ="SELECT dbo_ProductionProcessData.*, dbo_CableUnion.Location FROM dbo_CableUnion "
        "INNER JOIN dbo_ProductionProcessData ON dbo_CableUnion.Oracle = dbo_ProductionPr"
        "ocessData.[Item #];"
    Caption ="Item Search"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000ec6d0000e601000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
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
            Height =2040
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Oracle"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =16020
                    Top =1680
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label15"
                    Caption ="NCMIR #"
                    LayoutCachedLeft =16020
                    LayoutCachedTop =1680
                    LayoutCachedWidth =17400
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =11820
                    Top =1680
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Cut No"
                    LayoutCachedLeft =11820
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =2160
                    Top =1680
                    Width =3720
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Reel No"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =180
                    Top =420
                    Width =2760
                    Name ="Box35"
                End
                Begin Label
                    OverlapFlags =85
                    Left =12660
                    Top =1680
                    Width =2220
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label39"
                    Caption ="ID"
                    LayoutCachedLeft =12660
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =15000
                    Top =1680
                    Width =990
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="Buffer Line"
                    LayoutCachedLeft =15000
                    LayoutCachedTop =1680
                    LayoutCachedWidth =15990
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =8580
                    Top =1680
                    Width =2760
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="NCMIR #"
                    LayoutCachedLeft =8580
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =6000
                    Top =1680
                    Width =960
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Entry Date"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8520
                    Top =780
                    Width =1320
                    TabIndex =1
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =8520
                    LayoutCachedTop =780
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =300
                    Top =540
                    Width =2400
                    Height =480
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Process Data"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =7020
                    Top =780
                    Width =1320
                    TabIndex =2
                    Name ="btnSearch"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =780
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3180
                    Top =780
                    Width =3600
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3180
                    LayoutCachedTop =780
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7020
                    Top =1620
                    Width =480
                    TabIndex =3
                    Name ="btnSortDescending"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddd444d444dddd0dddd74ddd47ddd707dd ,
                        0xdd44444dddd000dddd74d47ddd70007dddd444dddd00000dddd747dddddd0ddd ,
                        0xdddd4ddddddd0ddddddddddddddd0dddd111111ddddd0dddd11ddd1ddddd0ddd ,
                        0xdd11dddddddd0dddddd11ddddddd0ddddddd11dddddd0dddd1ddd11ddddd0ddd ,
                        0xd111111ddddd0ddd000000000000000000000000000000000000000000000000 ,
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =7020
                    LayoutCachedTop =1620
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1980
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =3180
                    Top =60
                    Width =3900
                    Height =600
                    FontSize =9
                    Name ="Label65"
                    Caption ="Master Search: Searches all fields displayed with embedded wild card. "
                    LayoutCachedLeft =3180
                    LayoutCachedTop =60
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =660
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10020
                    Top =780
                    Width =480
                    TabIndex =4
                    Name ="Refresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dd0000ddd7227ddddd0ffd2d727727dd0f0ffd2727dd727d ,
                        0x0f0ffd227dddd72d0f0ffd2222dddddd0f0ffdddddd2222d0f0ffd27dddd722d ,
                        0x0f0ffd727dd7272d0f0fffd727727d2d0f0ffffd7227dddd0f0fffffd0000ddd ,
                        0x0f0ffffff0f08ddd0f0ffffff008dddd0f000000008ddddd0fffffffdddddddd ,
                        0x00000000dddddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =10020
                    LayoutCachedTop =780
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =17460
                    Top =1680
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label66"
                    Caption ="Location"
                    LayoutCachedLeft =17460
                    LayoutCachedTop =1680
                    LayoutCachedWidth =18300
                    LayoutCachedHeight =1980
                End
            End
        End
        Begin Section
            Height =426
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =16020
                    Top =60
                    Width =1380
                    Height =300
                    ColumnOrder =2
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="Text0"
                    ControlSource ="Fiber"

                    LayoutCachedLeft =16020
                    LayoutCachedTop =60
                    LayoutCachedWidth =17400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =1860
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="Text6"
                    ControlSource ="Item #"

                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =16020
                    Top =60
                    Width =1380
                    Height =300
                    ColumnOrder =5
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="Text16"
                    ControlSource ="NCMIR #"

                    LayoutCachedLeft =16020
                    LayoutCachedTop =60
                    LayoutCachedWidth =17400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11880
                    Top =60
                    Width =780
                    Height =300
                    ColumnWidth =2835
                    ColumnOrder =6
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="Text18"
                    ControlSource ="Cut No"

                    LayoutCachedLeft =11880
                    LayoutCachedTop =60
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2220
                    Top =60
                    Width =3720
                    Height =300
                    ColumnWidth =2280
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="Text22"
                    ControlSource ="Reel No"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =60
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =11400
                    Width =426
                    Height =426
                    Name ="btnFind"
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="OpenIcon.bmp"

                    LayoutCachedLeft =11400
                    LayoutCachedWidth =11826
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12720
                    Top =60
                    Width =2220
                    Height =300
                    ColumnWidth =2130
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text40"
                    ControlSource ="ID"

                    LayoutCachedLeft =12720
                    LayoutCachedTop =60
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15000
                    Top =60
                    Width =900
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="Text42"
                    ControlSource ="Buffer Line"

                    LayoutCachedLeft =15000
                    LayoutCachedTop =60
                    LayoutCachedWidth =15900
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =8580
                    Top =60
                    Width =2820
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="Text51"
                    ControlSource ="NCMIR #"

                    LayoutCachedLeft =8580
                    LayoutCachedTop =60
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =60
                    Width =2400
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="Text53"
                    ControlSource ="Entry Date"

                    LayoutCachedLeft =6060
                    LayoutCachedTop =60
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =17460
                    Top =60
                    Width =840
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="Text67"
                    ControlSource ="Fiber"

                    LayoutCachedLeft =17460
                    LayoutCachedTop =60
                    LayoutCachedWidth =18300
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =17460
                    Top =60
                    Width =840
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="Text68"
                    ControlSource ="Location"

                    LayoutCachedLeft =17460
                    LayoutCachedTop =60
                    LayoutCachedWidth =18300
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =1500
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =660
                    Width =1320
                    Name ="btnFiberClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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
'Key Preview = Yes


Private Sub btnFiberClose_Click()
DoCmd.Close acForm, Me.name

End Sub

Private Sub btnFind_Click()
Call OpenAClient(0, Me.[Item #], Nz(Me.Location, 2))
End Sub

Private Sub btnRemoveFilter_Click()
Me.ItemNoSearch.value = Null
Me.ItemNoSearch.Requery
Me.Filter = ""
Me.OrderBy = ""
Me.FilterOn = False
Me.OrderByOn = False

'Call Refresh_Click


End Sub




Private Sub btnSearch_Click()
Dim strFilter As String
strFilter = "[Item #] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Reel No] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Entry Date] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [NCMIR #] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Cut No] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [ID] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Buffer Line] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [NCMIR #] like """ & Me.ItemNoSearch & "*"""

Me.Filter = strFilter
Me.FilterOn = True
'Call Refresh_Click
End Sub



Private Sub btnSortDescending_Click()

Me.OrderBy = "[Entry Date] DESC"
Me.OrderByOn = True

'Call Refresh_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyW Then
        Call btnFiberClose_Click
    End If
End Sub


Private Sub ItemNoSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Me.btnSearch.SetFocus
        Call btnSearch_Click
    End If

End Sub


Private Sub Refresh_Click()
Me.Requery
'Me.Refresh
End Sub

Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =21960
    DatasheetFontHeight =10
    ItemSuffix =73
    Left =3135
    Top =1440
    Right =25095
    Bottom =11685
    DatasheetGridlinesColor =12632256
    Filter ="([tbl_Fibers].[Fiber_Type]=\"d\")"
    OrderBy ="[tbl_Fibers].[Fiber]"
    RecSrcDt = Begin
        0x4c7645be14c0e440
    End
    RecordSource ="tbl_Fibers"
    Caption ="Item Search"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000ec6d0000e601000001000000 ,
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
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Oracle"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =15000
                    Top =1680
                    Width =1740
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label15"
                    Caption ="Oracle Fiber Family"
                    LayoutCachedLeft =15000
                    LayoutCachedTop =1680
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =9600
                    Top =1680
                    Width =900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="OM#"
                    LayoutCachedLeft =9600
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =1920
                    Top =1680
                    Width =6900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Catalog Number"
                    LayoutCachedLeft =1920
                    LayoutCachedTop =1680
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =8880
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label24"
                    Caption ="Active"
                    LayoutCachedLeft =8880
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =16800
                    Top =1680
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label33"
                    Caption ="Vendor"
                    LayoutCachedLeft =16800
                    LayoutCachedTop =1680
                    LayoutCachedWidth =18480
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
                    Left =11580
                    Top =1680
                    Width =2220
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label39"
                    Caption ="Special Note"
                    LayoutCachedLeft =11580
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =13860
                    Top =1680
                    Width =1065
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="Fiber Letter"
                    LayoutCachedLeft =13860
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14925
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =10560
                    Top =1680
                    Width =900
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="BIF"
                    LayoutCachedLeft =10560
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =1980
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =1260
                    Width =21720
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =21900
                    LayoutCachedHeight =1260
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
                    Caption ="Item Search"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    Left =19680
                    Top =1680
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label61"
                    Caption ="ITU"
                    LayoutCachedLeft =19680
                    LayoutCachedTop =1680
                    LayoutCachedWidth =20760
                    LayoutCachedHeight =1980
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
                Begin Label
                    OverlapFlags =85
                    Left =3180
                    Top =60
                    Width =3900
                    Height =600
                    FontSize =9
                    Name ="Label65"
                    Caption ="Master Search: Searches all fields displayed with embedded wild card. Ex: PZ0028"
                        "0-01 or ZP002-2.0 or Molex"
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
                    TabIndex =3
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
                    Left =20820
                    Top =1680
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label66"
                    Caption ="Cost"
                    LayoutCachedLeft =20820
                    LayoutCachedTop =1680
                    LayoutCachedWidth =21900
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =20520
                    Top =660
                    Width =1320
                    Height =480
                    TabIndex =4
                    Name ="btnFiberCost"
                    Caption ="Update Fiber Cost"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =20520
                    LayoutCachedTop =660
                    LayoutCachedWidth =21840
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =18540
                    Top =1680
                    Width =1080
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label71"
                    Caption ="Tb Type"
                    LayoutCachedLeft =18540
                    LayoutCachedTop =1680
                    LayoutCachedWidth =19620
                    LayoutCachedHeight =1980
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1620
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =500
                    Name ="Text6"
                    ControlSource ="Fiber"

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15000
                    Top =60
                    Width =1740
                    Height =300
                    ColumnOrder =5
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="Text16"
                    ControlSource ="OracleFiberFamily"

                    LayoutCachedLeft =15000
                    LayoutCachedTop =60
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9600
                    Top =60
                    Width =900
                    Height =300
                    ColumnWidth =2835
                    ColumnOrder =6
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="Text18"
                    ControlSource ="OM#"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =6900
                    Height =300
                    ColumnWidth =2280
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="Text22"
                    ControlSource ="Description_Oracle"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8880
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="Text25"
                    ControlSource ="Active"
                    Format ="Yes/No"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =60
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16800
                    Top =60
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="Text34"
                    ControlSource ="Vendor"

                    LayoutCachedLeft =16800
                    LayoutCachedTop =60
                    LayoutCachedWidth =18480
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11580
                    Top =60
                    Width =2220
                    Height =300
                    ColumnWidth =2130
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="Text40"
                    ControlSource ="SpecialNotes"

                    LayoutCachedLeft =11580
                    LayoutCachedTop =60
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13860
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text42"
                    ControlSource ="Fiber_Type"

                    LayoutCachedLeft =13860
                    LayoutCachedTop =60
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10560
                    Top =60
                    Width =900
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="Text51"
                    ControlSource ="BIF"

                    LayoutCachedLeft =10560
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19680
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =10
                    Name ="Text62"
                    ControlSource ="ITU"

                    LayoutCachedLeft =19680
                    LayoutCachedTop =60
                    LayoutCachedWidth =20760
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20820
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="Text67"
                    ControlSource ="Cost"

                    LayoutCachedLeft =20820
                    LayoutCachedTop =60
                    LayoutCachedWidth =21900
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18540
                    Top =60
                    Width =1080
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="Text72"
                    ControlSource ="TBCatalogType"

                    LayoutCachedLeft =18540
                    LayoutCachedTop =60
                    LayoutCachedWidth =19620
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


Private Sub btnFiberClose_Click()
DoCmd.Close acForm, Me.name

End Sub



Private Sub btnFiberCost_Click()
DoCmd.OpenQuery "qryUPDATEqryFibersCost"
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
strFilter = "[fiber] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Description_Oracle] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [OM#] like """ & Me.[ItemNoSearch] & "*"""
strFilter = strFilter & " OR [SpecialNotes] like """ & Me.ItemNoSearch & "*"""
strFilter = strFilter & " OR [Vendor] like """ & Me.ItemNoSearch & "*"""

Me.Filter = strFilter
Me.FilterOn = True
'Call Refresh_Click
End Sub



Private Sub btnSortDescending_Click()

Me.OrderBy = "[RevDate] DESC"
Me.OrderByOn = True

'Call Refresh_Click

End Sub



Private Sub Form_Load()
Call UserCheck
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

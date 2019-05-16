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
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15000
    DatasheetFontHeight =10
    ItemSuffix =81
    Left =6615
    Top =3105
    Right =21615
    Bottom =8370
    DatasheetGridlinesColor =12632256
    Filter ="[Unit Color Series] = 'PG08374'"
    OrderBy ="[frmFindUnitSeries].[UnitNumber], [frmFindUnitSeries].[Unit ID] DESC"
    RecSrcDt = Begin
        0x9567fcde2adfe440
    End
    RecordSource ="SELECT [Unit Series].*, [Basic Product Construction].[Item No] FROM [Unit Series"
        "] LEFT JOIN [Basic Product Construction] ON [Unit Series].OracleItem = [Basic Pr"
        "oduct Construction].[New Oracle Part #];"
    Caption ="Item Search"
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
            Height =1980
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
                    Caption ="Unit Color Series"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =4260
                    Top =1680
                    Width =3480
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Unit Type"
                    LayoutCachedLeft =4260
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7740
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
                    Left =9240
                    Top =1680
                    Width =990
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label41"
                    Caption ="Print Type"
                    LayoutCachedLeft =9240
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10230
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =7800
                    Top =1680
                    Width =1410
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label49"
                    Caption ="TB Color Series"
                    LayoutCachedLeft =7800
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1260
                    Width =14040
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =1260
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
                    Caption ="Unit Series"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    Left =11340
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label69"
                    Caption ="Fiber Code"
                    LayoutCachedLeft =11340
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =12420
                    Top =1680
                    Width =990
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label73"
                    Caption ="ID"
                    LayoutCachedLeft =12420
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13410
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3240
                    Top =780
                    Width =1320
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =780
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4860
                    Top =780
                    Width =1320
                    TabIndex =1
                    Name ="btnAdd"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4860
                    LayoutCachedTop =780
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6480
                    Top =780
                    Width =1320
                    TabIndex =2
                    Name ="btnSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6480
                    LayoutCachedTop =780
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =1140
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =2940
                    Top =1680
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label78"
                    Caption ="Oracle Item"
                    LayoutCachedLeft =2940
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =1920
                    Top =1680
                    Width =960
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label80"
                    Caption ="Unit #"
                    LayoutCachedLeft =1920
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1980
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12240
                    Top =480
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txUnitSeries"

                    LayoutCachedLeft =12240
                    LayoutCachedTop =480
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =780
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
                    Left =240
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="Base"
                    ControlSource ="Unit Color Series"

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
                    Left =4260
                    Top =60
                    Width =3480
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txItemNo"
                    ControlSource ="Item No"

                    LayoutCachedLeft =4260
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9240
                    Top =60
                    Width =2040
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="Comments"
                    ControlSource ="Print type"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =60
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7800
                    Top =60
                    Width =1380
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="ListingSection"
                    ControlSource ="TB Color Series"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11340
                    Top =60
                    Width =1020
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text72"
                    ControlSource ="Fiber Code"

                    LayoutCachedLeft =11340
                    LayoutCachedTop =60
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12420
                    Top =60
                    Width =1020
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="Text74"
                    ControlSource ="Number"

                    LayoutCachedLeft =12420
                    LayoutCachedTop =60
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13680
                    Top =60
                    Width =480
                    Height =300
                    TabIndex =8
                    Name ="btnDelete"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddddddddddddddddddddddddddddddd0fd ,
                        0xddd0fddddddddddddd000fdddddd0fdddd000fddddd0fdddddd000fddd00fddd ,
                        0xdddd000fd00fddddddddd00000fddddddddddd000fddddddddddd00000fddddd ,
                        0xdddd000fd0fddddddd0000fddd00fdddd0000fddddd00fddd00fdddddddd00fd ,
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =13680
                    LayoutCachedTop =60
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2940
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txOracleItem"
                    ControlSource ="OracleItem"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2940
                    LayoutCachedTop =60
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =960
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="Text81"
                    ControlSource ="UnitNumber"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =2880
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
                    OverlapFlags =85
                    Left =300
                    Top =660
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1920
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="Refresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

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
Option Explicit
Option Compare Database
Dim m_unitSeries


Friend Property Get UnitSeries() As String
If IsNull(m_unitSeries) Or m_unitSeries = "" Then
    m_unitSeries = getUnitSeries
End If
UnitSeries = m_unitSeries
End Property


Private Sub btnAdd_Click()
Me.AllowAdditions = True
Call btnEdit_Click
End Sub

Private Sub btnClose_Click()
On Error GoTo ErrorHandler
CloseClient (Me.hWnd)

ErrorExit:
Exit Sub

ErrorHandler:
DoCmd.Close acForm, Me.name
Resume ErrorExit
End Sub

Private Sub btnDelete_Click()
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
Me.txItemNo.Locked = True
End Sub


Private Function getUnitSeries() As String
Dim hold$()
hold = Split(Me.Filter, "'")
getUnitSeries = hold(1)
End Function



Private Sub btnSave_Click()
Me.AllowAdditions = False
Me.Requery
End Sub



Private Sub Form_Open(Cancel As Integer)
Me.AllowAdditions = False
getUnitSeries
End Sub

Private Sub Refresh_Click()
Me.Requery
Me.Refresh
End Sub

Private Sub txOracleItem_AfterUpdate()
If Me.FilterOn = True And UnitSeriesExist(UnitSeries) Then
    Me.[Unit Color Series] = UnitSeries
End If
End Sub

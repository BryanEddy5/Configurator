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
    Width =11340
    DatasheetFontHeight =10
    ItemSuffix =87
    Left =8445
    Top =2760
    Right =19785
    Bottom =9045
    DatasheetGridlinesColor =12632256
    Filter ="([frmWhereUsed].[OracleStatus]=\"Active\")"
    OrderBy ="[frmWhereUsed].[Item No], [frmWhereUsed].[AssemblyItemNumber]"
    RecSrcDt = Begin
        0xd1d884d92afee440
    End
    RecordSource ="SELECT tblWhereUsed.ComponentItemNumber, tblWhereUsed.AssemblyItemNumber, tblWhe"
        "reUsed.Quantity, [Basic Product Construction].OracleStatus, [Basic Product Const"
        "ruction].[Item No], tblCableConstructions.CablePasses FROM (([Basic Product Cons"
        "truction] INNER JOIN tblWhereUsed ON [Basic Product Construction].[New Oracle Pa"
        "rt #] = tblWhereUsed.AssemblyItemNumber) INNER JOIN tblCableConstructionReferenc"
        "es ON [Basic Product Construction].Base = tblCableConstructionReferences.Base) I"
        "NNER JOIN tblCableConstructions ON tblCableConstructionReferences.BaseID = tblCa"
        "bleConstructions.BaseID ORDER BY tblWhereUsed.ComponentItemNumber; "
    Caption ="Item Search"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyPress ="[Event Procedure]"
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
            Height =2820
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =2520
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Item"
                    LayoutCachedLeft =780
                    LayoutCachedTop =2520
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =2880
                    Top =2520
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Assembly"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =8400
                    Top =2520
                    Width =960
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Quantity"
                    LayoutCachedLeft =8400
                    LayoutCachedTop =2520
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2820
                End
                Begin Line
                    OverlapFlags =85
                    Top =1440
                    Width =10500
                    Name ="Line58"
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1440
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =3480
                    Top =1740
                    Width =3420
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItem"
                    Caption ="RBN00147"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3480
                    LayoutCachedTop =1740
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =2235
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =3000
                    Top =300
                    Width =4740
                    Name ="Box35"
                    LayoutCachedLeft =3000
                    LayoutCachedTop =300
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =3120
                    Top =420
                    Width =4500
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Item Where Used"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3120
                    LayoutCachedTop =420
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =915
                End
                Begin Label
                    OverlapFlags =85
                    Left =9480
                    Top =2520
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label81"
                    Caption ="Status"
                    LayoutCachedLeft =9480
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =4560
                    Top =2520
                    Width =3720
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label83"
                    Caption ="Catalog"
                    LayoutCachedLeft =4560
                    LayoutCachedTop =2520
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =2820
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
                    Left =780
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="Base"
                    ControlSource ="ComponentItemNumber"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =780
                    LayoutCachedTop =60
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =1560
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="ListingCompany"
                    ControlSource ="AssemblyItemNumber"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8400
                    Top =60
                    Width =975
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="ListingTab"
                    ControlSource ="Quantity"

                    LayoutCachedLeft =8400
                    LayoutCachedTop =60
                    LayoutCachedWidth =9375
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9480
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="Text82"
                    ControlSource ="OracleStatus"

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =2400
                    Width =426
                    Height =426
                    TabIndex =4
                    Name ="btnFind"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadad000dadada ,
                        0xadadad0bb00dadaddadada0bb00adadaadadada00dadadaddadada0b00dadada ,
                        0xadadad0b00adadaddadada0b00dadadaadadad0bb00dadaddada00d0bb00dada ,
                        0xada0b00d0bb00daddad0b00ad0b00adaada0bb000bb00daddada0bbbbb00dada ,
                        0xadada000000dadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =2400
                    LayoutCachedWidth =2826
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
                    Left =4560
                    Top =60
                    Width =3720
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="Text84"
                    ControlSource ="Item No"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =60
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10860
                    Top =60
                    Width =480
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text86"
                    ControlSource ="CablePasses"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =60
                    LayoutCachedWidth =11340
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


Private Sub btnFind_Click()
Call OpenAClient(0, Me.AssemblyItemNumber, Nz(Me.CablePasses, 2))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

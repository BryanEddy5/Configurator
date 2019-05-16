Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5280
    RowHeight =255
    DatasheetFontHeight =10
    ItemSuffix =97
    Left =9990
    Top =1245
    Right =15780
    Bottom =7245
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2e027770abf3e440
    End
    RecordSource ="tblMaterialBOMExplode"
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
            Height =1980
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =1440
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Item"
                    LayoutCachedLeft =1440
                    LayoutCachedTop =1680
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Top =1200
                    Width =5100
                    Name ="Line58"
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =1200
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =120
                    Top =300
                    Width =4740
                    Name ="Box35"
                    LayoutCachedLeft =120
                    LayoutCachedTop =300
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =210
                    Top =420
                    Width =4545
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Material Where Used"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =210
                    LayoutCachedTop =420
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =915
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4500
                    Width =366
                    Height =306
                    ForeColor =0
                    Name ="cmdDelete"
                    Caption ="Command25"
                    OnClick ="mcrDeleteRecord"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =4500
                    LayoutCachedWidth =4866
                    LayoutCachedHeight =306
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =120
                    Height =255
                    TabIndex =1
                    Name ="Text95"
                    ControlSource ="OracleItem"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =120
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =375
                End
            End
        End
        Begin FormFooter
            Height =840
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =180
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =300
                    LayoutCachedTop =180
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =540
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



Private Sub btnExcelExport_Click()

On Error GoTo ErrorHandler

DoCmd.OutputTo acOutputForm, "frmGetBuyMaterials", acFormatXLS, Environ("USERPROFILE") & "\Desktop\BuyMaterials" & Format(Now, "yyyyMMddhhmmss") & ".xls"
Beep
MsgBox "File has successfully been exported to Desktop"

ErrorExit:
Exit Sub

ErrorHandler:
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

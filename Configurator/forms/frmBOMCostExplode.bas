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
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13260
    DatasheetFontHeight =10
    ItemSuffix =92
    Left =-21315
    Top =1320
    Right =-8055
    Bottom =11925
    DatasheetGridlinesColor =12632256
    Filter ="([frmBOMCostExplode].[ComponentItemNumber] Like \"bin*\")"
    OrderBy ="[frmBOMCostExplode].[ComponentItemNumber]"
    RecSrcDt = Begin
        0x303f4a9793ece440
    End
    RecordSource ="SELECT tblBOMComponentExplode.BOMLevel, tblBOMComponentExplode.AssemblyItemNumbe"
        "r, tblBOMComponentExplode.Weight_kg_km, tblBOMComponentExplode.ItemCost_km, tblB"
        "OMComponentExplode.Cost_Ratio, tblBOMComponentExplode.FinishedGood, tblBOMCompon"
        "entExplode.ComponentItemNumber, tblBOMComponentExplode.Total_Cost_km FROM tblBOM"
        "ComponentExplode ORDER BY tblBOMComponentExplode.BOMLevel, tblBOMComponentExplod"
        "e.AssemblyItemNumber, tblBOMComponentExplode.ItemCost_km;"
    Caption ="Item Search"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyPress ="[Event Procedure]"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
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
                    Left =3720
                    Top =2520
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="txMaterial"
                    Caption ="Material"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =2520
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =5400
                    Top =2520
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="BOM Level"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =2520
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =7020
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Cost ($/km)"
                    LayoutCachedLeft =7020
                    LayoutCachedTop =2520
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =2820
                End
                Begin Line
                    OverlapFlags =85
                    Left =240
                    Top =1140
                    Width =13020
                    Name ="Line58"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1140
                    LayoutCachedWidth =13260
                    LayoutCachedHeight =1140
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =3840
                    Top =300
                    Width =4740
                    Name ="Box35"
                    LayoutCachedLeft =3840
                    LayoutCachedTop =300
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =3960
                    Top =420
                    Width =4500
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="BOM Component Explode"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3960
                    LayoutCachedTop =420
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =915
                End
                Begin Label
                    OverlapFlags =85
                    Left =2040
                    Top =2520
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="txMake_Item"
                    Caption ="Make Item"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =2520
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =2520
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label83"
                    Caption ="Finished Good"
                    LayoutCachedLeft =360
                    LayoutCachedTop =2520
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =10020
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label85"
                    Caption ="Weight (kg/km)"
                    LayoutCachedLeft =10020
                    LayoutCachedTop =2520
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =8520
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label87"
                    Caption ="Cost Ratio"
                    LayoutCachedLeft =8520
                    LayoutCachedTop =2520
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    Left =11580
                    Top =2520
                    Width =1440
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label90"
                    Caption ="Total ($/km)"
                    LayoutCachedLeft =11580
                    LayoutCachedTop =2520
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =4620
                    Top =1260
                    Width =3420
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItem"
                    Caption ="PZ08192-10"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =4620
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =1755
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =3060
                    Top =1860
                    Width =6480
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblDesignCode"
                    Caption ="ZP002F202701"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =3060
                    LayoutCachedTop =1860
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =2355
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
                    Left =3720
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="txComponent"
                    ControlSource ="ComponentItemNumber"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5400
                    Top =60
                    Width =1560
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="txBOMLevel"
                    ControlSource ="BOMLevel"

                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="txCost_km"
                    ControlSource ="ItemCost_km"
                    Format ="General Number"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =60
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txMakeItem"
                    ControlSource ="AssemblyItemNumber"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10020
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="txWeight_kg_m"
                    ControlSource ="Weight_kg_km"

                    LayoutCachedLeft =10020
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="txFinishedGood"
                    ControlSource ="FinishedGood"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8520
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    Name ="txCostRatio"
                    ControlSource ="Cost_Ratio"
                    Format ="General Number"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11580
                    Top =60
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="txTotal_Cost_km"
                    ControlSource ="Total_Cost_km"
                    Format ="General Number"

                    LayoutCachedLeft =11580
                    LayoutCachedTop =60
                    LayoutCachedWidth =13020
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =1860
                    Top =300
                    Width =540
                    TabIndex =1
                    Name ="btnExcelExport"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddd87262ddddddd876262626dddddd6262626262262628 ,
                        0x2626262626fffff662ff6ff2626f62f2268f2f8626fffff6626fff62626f62f2 ,
                        0x2628f82626fffff66268f862626f62f2262fff2626fffff6628f6f82626f62f2 ,
                        0x26ff2ff626fffff662626262626f62f22626262626fffff6d872626262626268 ,
                        0xddddd87626dddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =1860
                    LayoutCachedTop =300
                    LayoutCachedWidth =2400
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

Private Sub txFinishedGood_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

Private Sub btnClose_Click()
DoCmd.Close acForm, Me.name
End Sub


Private Sub btnExcelExport_Click()
On Error GoTo ErrorHandler

DoCmd.OutputTo acOutputForm, "frmBOMCostExplode", acFormatXLS, Environ("USERPROFILE") & "\Desktop\BOMCostExplode" & Format(Now, "yyyyMMddhhmmss") & ".xls"
Beep
MsgBox "File has successfully been exported to Desktop"



ErrorExit:
Exit Sub

ErrorHandler:
If Err.Number = 2302 Then
    MsgBox "Close the excel spreadsheet and try again", vbCritical
Else
    MsgBox Err.Number & " " & Err.Description
End If
Resume ErrorExit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
DoCmd.Close acForm, Me.name
End If
End Sub

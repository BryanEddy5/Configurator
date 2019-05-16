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
    FilterOn = NotDefault
    OrderByOn = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =16125
    DatasheetFontHeight =10
    ItemSuffix =73
    Left =6060
    Top =825
    Right =22185
    Bottom =12930
    DatasheetGridlinesColor =12632256
    Filter ="[ItemNo] like \"*\""
    OrderBy ="[dbo_tblStockItems].[StockOracleItem], [dbo_tblStockItems].[ItemNo], [dbo_tblSto"
        "ckItems].[Price_feet] DESC, [dbo_tblStockItems].[Min_Order_Quantity]"
    RecSrcDt = Begin
        0xaa8741ee51ece440
    End
    RecordSource ="dbo_tblStockItems"
    Caption ="Stock Pricing"
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
                    Left =3540
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Price ($/ft)"
                    LayoutCachedLeft =3540
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =1680
                    Width =3300
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Catalog"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =240
                    Top =60
                    Width =2940
                    Name ="Box35"
                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =780
                End
                Begin Label
                    OverlapFlags =85
                    Left =5460
                    Top =1680
                    Width =960
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="MOQ"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1380
                    Width =11580
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1380
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =1380
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5520
                    Top =900
                    Width =1320
                    TabIndex =1
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =900
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1260
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =360
                    Top =180
                    Width =2700
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Stock Pricing"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =360
                    LayoutCachedTop =180
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =675
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =4020
                    Top =900
                    Width =1320
                    TabIndex =2
                    Name ="btnFilter"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =900
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1260
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
                    Left =180
                    Top =900
                    Width =3600
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    Name ="ItemNoSearch"
                    Format ="@;\"<enter search text>\""
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =900
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =85
                    Left =6480
                    Top =1680
                    Width =4800
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label68"
                    Caption ="Lead Time Name"
                    LayoutCachedLeft =6480
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7080
                    Top =900
                    Width =1320
                    TabIndex =3
                    Name ="Refresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =900
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =1260
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =11340
                    Top =1680
                    Width =3060
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label72"
                    Caption ="Oracle Item Number"
                    LayoutCachedLeft =11340
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =1980
                End
            End
        End
        Begin Section
            Height =366
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3540
                    Top =60
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="CusomerSpec"
                    ControlSource ="Price_feet"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =60
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =3300
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="CustomerPartNumber"
                    ControlSource ="ItemNo"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5460
                    Top =60
                    Width =960
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="Customer"
                    ControlSource ="Min_Order_Quantity"

                    LayoutCachedLeft =5460
                    LayoutCachedTop =60
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14880
                    Width =366
                    Height =366
                    TabIndex =4
                    Name ="cmdDelete"
                    Caption ="Command12"
                    OnClick ="[Event Procedure]"
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"

                    LayoutCachedLeft =14880
                    LayoutCachedWidth =15246
                    LayoutCachedHeight =366
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6480
                    Top =60
                    Width =4785
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="txLeadTime"
                    ControlSource ="LeadTime_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesForce_vLeadTimes.LeadTimeID, SalesForce_vLeadTimes.LeadTimeName FROM"
                        " SalesForce_vLeadTimes;"
                    ColumnWidths ="0;4320"
                    LayoutCachedLeft =6480
                    LayoutCachedTop =60
                    LayoutCachedWidth =11265
                    LayoutCachedHeight =360
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11340
                    Top =60
                    Width =3360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="cbOracleItem"
                    ControlSource ="StockOracleItem"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW [Basic Product Construction].[New Oracle Part #] FROM [Basic "
                        "Product Construction] ORDER BY [Basic Product Construction].[New Oracle Part #];"
                    ColumnWidths ="2880"

                    LayoutCachedLeft =11340
                    LayoutCachedTop =60
                    LayoutCachedWidth =14700
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
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =660
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3420
                    Top =660
                    Width =1500
                    TabIndex =2
                    Name ="btnAddRecord"
                    Caption ="Add New Record"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =660
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5160
                    Top =660
                    Width =1500
                    TabIndex =3
                    Name ="btnSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5160
                    LayoutCachedTop =660
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =1020
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

Private Sub btnAddRecord_Click()
Me.AllowAdditions = True
DoCmd.GoToRecord , , acNewRec
Call btnEdit_Click
End Sub


Private Sub btnClose_Click()
DoCmd.Close
End Sub

Private Sub LockControls(LockFields As Boolean)
Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
        ctl.Locked = LockFields
    End If
Next ctl

End Sub

Private Sub btnEdit_Click()
LockControls (False)
End Sub

Private Sub btnFilter_Click()

Dim strFilter As String
strFilter = "[ItemNo] like """ & Me.ItemNoSearch & "*"""


Me.Filter = strFilter
Me.FilterOn = True
Call Refresh_Click


End Sub

Private Sub btnRemoveFilter_Click()
Me.ItemNoSearch.value = Null
Me.ItemNoSearch.Requery
Me.Filter = ""
Me.FilterOn = False
Call Refresh_Click
End Sub

Private Sub btnSave_Click()
LockControls (False)
Me.AllowAdditions = False
FilterForm

Refresh_Click
End Sub

Private Sub cmdDelete_Click()
   If MsgBox("Are you SURE you want to delete this record?" & vbCrLf & _
    "This will permanently delete the record.", vbYesNo, "2nd Delete Confirmation") = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE FROM dbo_tblStockItems WHERE ItemNo = '" & Me.ItemNo & "'"
        DoCmd.SetWarnings True
        Me.Requery
   End If
End Sub

Private Sub FilterForm()
Dim strFilter As String
strFilter = "[ItemNo] like """ & Me.ItemNoSearch & "*"""

Me.Filter = strFilter
Me.FilterOn = True
End Sub


Private Sub Form_Open(Cancel As Integer)
LockControls (True)
FilterForm
Call Refresh_Click


End Sub

Private Sub ItemNoSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Me.btnFilter.SetFocus
        Call btnFilter_Click
    End If
End Sub


Private Sub Refresh_Click()
Me.Requery
Me.Refresh
End Sub

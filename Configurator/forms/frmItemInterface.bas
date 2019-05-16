Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =28020
    DatasheetFontHeight =10
    ItemSuffix =65
    Top =60
    Right =22830
    Bottom =13455
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xae8c47ded6cee440
    End
    RecordSource ="qryItemInterface"
    Caption ="Oracle Item Interface Tables"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =900
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =300
                    Top =600
                    Width =480
                    Height =300
                    FontSize =10
                    Name ="Label14"
                    Caption ="Org"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =840
                    Top =600
                    Width =1080
                    Height =300
                    FontSize =10
                    Name ="Label15"
                    Caption ="Item Number"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2700
                    Top =600
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label16"
                    Caption ="Description"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =5760
                    Top =600
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="UOM"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =6480
                    Top =600
                    Width =1680
                    Height =300
                    FontSize =10
                    Name ="Label18"
                    Caption ="Template Name"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =9360
                    Top =600
                    Width =2100
                    Height =300
                    FontSize =10
                    Name ="Label19"
                    Caption ="Catalog Group Name"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =11940
                    Top =600
                    Width =1140
                    Height =300
                    FontSize =10
                    Name ="Label20"
                    Caption ="Std Lot Size"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =12360
                    Top =60
                    Width =900
                    FontSize =10
                    FontWeight =700
                    Name ="cmdLoad"
                    Caption ="Load"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =13380
                    Top =60
                    Width =900
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =13380
                    LayoutCachedTop =60
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =15900
                    Top =540
                    Width =660
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =2
                    BackColor =255
                    Name ="PrintItemNumber"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =15900
                    LayoutCachedTop =540
                    LayoutCachedWidth =16560
                    LayoutCachedHeight =840
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =15720
                    Top =600
                    ColumnOrder =1
                    TabIndex =3
                    Name ="PrintFlag"
                    DefaultValue ="-1"

                    LayoutCachedLeft =15720
                    LayoutCachedTop =600
                    LayoutCachedWidth =15980
                    LayoutCachedHeight =840
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =6540
                    Top =60
                    Width =1800
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="cmdPrint"
                    Caption ="Item/BOM/Routing"
                    OnClick ="mcrItemInterfacePrint"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =4680
                    Top =60
                    Width =1800
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="cmdSetup"
                    Caption ="Print Setup Info."
                    OnClick ="mcrItemInterfaceSetupPrint"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2820
                    Top =60
                    Width =1800
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="cmdItemAttributes"
                    Caption ="Load Attributes"
                    OnClick ="mcrItemAttributesSave"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =360
                    Top =60
                    Width =1680
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="cmdPrintExt"
                    Caption ="Add a Print Ext"
                    OnClick ="mcrNewPrintExtensionFromBaseOpen"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =10800
                    Top =60
                    Width =1380
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    Name ="cmdQTSOpen"
                    Caption ="QTS Specs"
                    OnClick ="mcrQTSOpen"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =13200
                    Top =600
                    Width =1170
                    Height =300
                    FontSize =10
                    Name ="Label61"
                    Caption ="Cable Weight"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =600
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =900
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =14460
                    Top =600
                    Width =1170
                    Height =300
                    FontSize =10
                    Name ="Label63"
                    Caption ="Cable Width"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =14460
                    LayoutCachedTop =600
                    LayoutCachedWidth =15630
                    LayoutCachedHeight =900
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =300
                    Top =60
                    Width =480
                    Height =300
                    ColumnWidth =465
                    FontSize =10
                    Name ="OrganizationCode"
                    ControlSource ="OrganizationCode"
                    OnEnter ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =780
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =840
                    Top =60
                    Width =1800
                    Height =300
                    ColumnWidth =1185
                    FontSize =10
                    TabIndex =1
                    Name ="ItemNumber"
                    ControlSource ="ItemNumber"
                    OnEnter ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =2700
                    Top =60
                    Width =3000
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"
                    OnEnter ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =60
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =5760
                    Top =60
                    Width =600
                    Height =300
                    ColumnWidth =465
                    FontSize =10
                    TabIndex =3
                    Name ="PrimaryUOMCode"
                    ControlSource ="PrimaryUOMCode"
                    OnEnter ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =6420
                    Top =60
                    Width =2880
                    Height =300
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =4
                    Name ="TemplateName"
                    ControlSource ="TemplateName"
                    OnEnter ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6420
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =9360
                    Top =60
                    Width =2520
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =5
                    Name ="CatalogGroupName"
                    ControlSource ="CatalogGroupName"
                    OnEnter ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =60
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =11940
                    Top =60
                    Width =1200
                    Height =300
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =6
                    Name ="StdLotSize"
                    ControlSource ="StdLotSize"
                    OnEnter ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =11940
                    LayoutCachedTop =60
                    LayoutCachedWidth =13140
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =13200
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="Text62"
                    ControlSource ="CableWeight"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =13200
                    LayoutCachedTop =60
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =14460
                    Top =60
                    Width =1200
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="Text64"
                    ControlSource ="CableWidth"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =14460
                    LayoutCachedTop =60
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =9480
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =8520
                    Top =120
                    Width =7848
                    Height =7080
                    Name ="frmItemInterfaceCatalog"
                    SourceObject ="Form.frmItemInterfaceCatalog"
                    LinkChildFields ="ItemNumber"
                    LinkMasterFields ="ItemNumber"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =120
                    LayoutCachedWidth =16368
                    LayoutCachedHeight =7200
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =240
                    Top =120
                    Width =7416
                    Height =7080
                    TabIndex =1
                    Name ="frmItemInterfaceCategory"
                    SourceObject ="Form.frmItemInterfaceCategory"
                    LinkChildFields ="OrganizationCode;ItemNumber"
                    LinkMasterFields ="OrganizationCode;ItemNumber"

                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =7656
                    LayoutCachedHeight =7200
                End
                Begin Subform
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =9888
                    Top =120
                    Width =6480
                    Height =6780
                    TabIndex =2
                    Name ="frmItemInterfaceRouting"
                    SourceObject ="Form.frmItemInterfaceRouting"
                    LinkChildFields ="OrganizationCode"
                    LinkMasterFields ="OrganizationCode"

                    LayoutCachedLeft =9888
                    LayoutCachedTop =120
                    LayoutCachedWidth =16368
                    LayoutCachedHeight =6900
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =11148
                    Top =8580
                    Width =960
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdWOP"
                    Caption ="WOP"
                    OnClick ="mcrWOPOpen"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =11148
                    LayoutCachedTop =8580
                    LayoutCachedWidth =12108
                    LayoutCachedHeight =8940
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =13308
                    Top =8580
                    Width =1980
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =1279872587
                    Name ="cmdOracleInstructions"
                    Caption ="Oracle Instructions"
                    FontName ="Times New Roman"
                    HyperlinkAddress ="F:\\Temporary\\Kim McConnell\\Instructional"

                    LayoutCachedLeft =13308
                    LayoutCachedTop =8580
                    LayoutCachedWidth =15288
                    LayoutCachedHeight =8940
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =7320
                    Width =7416
                    Height =1020
                    TabIndex =5
                    Name ="frmItemInterfaceBOMBillOfMtlsInterface"
                    SourceObject ="Form.frmItemInterfaceBOMBillOfMtlsInterface"
                    LinkChildFields ="OrganizationCode;ItemNumber"
                    LinkMasterFields ="OrganizationCode;ItemNumber"

                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =8628
                    Top =7320
                    Width =7416
                    Height =1020
                    TabIndex =6
                    Name ="frmItemInterfaceRoutingInterface"
                    SourceObject ="Form.frmItemInterfaceRoutingInterface"
                    LinkChildFields ="OrganizationCode;ItemNumber"
                    LinkMasterFields ="OrganizationCode;ItemNumber"

                    LayoutCachedLeft =8628
                    LayoutCachedTop =7320
                    LayoutCachedWidth =16044
                    LayoutCachedHeight =8340
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =16560
                    Top =120
                    Width =6840
                    Height =3120
                    TabIndex =7
                    Name ="frmPremiseItemDetails"
                    SourceObject ="Form.frmPremiseItemDetails"

                    LayoutCachedLeft =16560
                    LayoutCachedTop =120
                    LayoutCachedWidth =23400
                    LayoutCachedHeight =3240
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =16560
                    Top =3300
                    Width =6840
                    Height =2880
                    TabIndex =8
                    Name ="frmPremiseItemDetailsPrintExt"
                    SourceObject ="Form.frmPremiseItemDetailsPrintExt"

                    LayoutCachedLeft =16560
                    LayoutCachedTop =3300
                    LayoutCachedWidth =23400
                    LayoutCachedHeight =6180
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =16560
                    Top =6240
                    Width =6840
                    Height =3000
                    TabIndex =9
                    Name ="frmPremiseItemDetailsCopied"
                    SourceObject ="Form.frmPremiseItemDetailsCopied"

                    LayoutCachedLeft =16560
                    LayoutCachedTop =6240
                    LayoutCachedWidth =23400
                    LayoutCachedHeight =9240
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =240
                    Top =120
                    Width =9660
                    Height =6720
                    TabIndex =10
                    Name ="frmItemInterfaceBOM"
                    SourceObject ="Form.frmItemInterfaceBOM"
                    LinkChildFields ="OrganizationCode"
                    LinkMasterFields ="OrganizationCode"

                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =6840
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

Private Sub CatalogGroupName_Enter()

    Call sItemInterfaceDisplay
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.name
    Forms.frmMainMenu.Visible = True
    DoCmd.Close acForm, "frmMainMenu"

End Sub

Private Sub cmdLoad_Click()

    DoCmd.OpenQuery "qryCategoryBlanksDelete"
    Me!frmItemInterfaceCategory.Form.Requery
    Me!frmItemInterfaceCategory.Form.Refresh


    If FieldCompletionCheck = True Then
        Exit Sub
    End If


    Dim WSHShell
    Set WSHShell = VBA.CreateObject("WScript.Shell")
    Dim exeName
    Dim statusCode
    Dim theDir
    Dim OldDir

    OldDir = WSHShell.CurrentDirectory
    theDir = "C:\Applications\Prod\program_updates"
    WSHShell.CurrentDirectory = theDir
    exeName = "C:\Applications\prod\program_updates\qts_program_update.exe APPS XXX CDI"
    statusCode = WSHShell.Run(exeName, 1, True)
    WSHShell.CurrentDirectory = OldDir


End Sub

Private Sub Description_Enter()
    Call sItemInterfaceDisplay
End Sub

Private Sub ItemNumber_Enter()
    Call sItemInterfaceDisplay
End Sub

Private Sub OrganizationCode_Enter()
    Call sItemInterfaceDisplay
End Sub

Private Sub PrimaryUOMCode_Enter()
    Call sItemInterfaceDisplay
End Sub

Private Sub StdLotSize_Enter()
    Call sItemInterfaceDisplay
End Sub

Private Sub TemplateName_Enter()
    Call sItemInterfaceDisplay
End Sub

Function FieldCompletionCheck() As Boolean
    If IsNull(Me!frmItemInterfaceRoutingInterface.Form.CompletionSubinventory) = True Then
        MsgBox "Completion subinventory is missing"
        FieldCompletionCheck = False
        Exit Function
    End If



End Function

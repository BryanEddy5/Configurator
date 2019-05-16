Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6060
    DatasheetFontHeight =10
    ItemSuffix =62
    Left =5670
    Top =8760
    Right =11730
    Bottom =11805
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x379d288a0cffe440
    End
    RecordSource ="qryPremiseItemSelected"
    Caption ="Enter New Item Number"
    OnOpen ="[Event Procedure]"
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
            Height =1560
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =240
                    Top =240
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label40"
                    Caption ="Select New Item"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1680
                    Top =1080
                    Width =2880
                    Height =300
                    ColumnOrder =1
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="LotSize"
                    Format =">"
                    AfterUpdate ="mcrNewItemNumberSetLotSize"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =240
                    Top =1080
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label42"
                    Caption ="Enter Lot Size"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =8640
                    Left =1680
                    Top =240
                    Width =2880
                    Height =300
                    ColumnWidth =3000
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    ColumnInfo ="\"Oracle Item Number\";\"\";\"Catalog Number\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="NewItemNumber"
                    RowSourceType ="Table/Query"
                    RowSource ="qryPremiseItemSelect"
                    ColumnWidths ="2160;2160;4320"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    Format =">"

                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =240
                    Top =660
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label55"
                    Caption ="Item Type"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1680
                    Top =660
                    Width =2880
                    Height =300
                    ColumnOrder =2
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="NewItemType"
                    RowSourceType ="Table/Query"
                    RowSource ="qryItemTypeSelect"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    Format =">"

                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =4620
                    Top =240
                    Width =540
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BackColor =255
                    Name ="ItemSuffix"
                    ControlSource ="=Right([NewItemNumber],2)"
                    AfterUpdate ="mcrNewItemNumberSetLotSize"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4620
                    LayoutCachedTop =240
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =4620
                    Top =660
                    Width =600
                    Height =300
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BackColor =255
                    Name ="BaseItem"
                    ControlSource ="=Left([NewItemNumber],7)"
                    AfterUpdate ="mcrNewItemNumberSetLotSize"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =5280
                    Top =660
                    Width =600
                    Height =300
                    ColumnOrder =5
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BackColor =255
                    Name ="ItemPrefix"
                    ControlSource ="=Left([NewItemNumber],3)"
                    AfterUpdate ="mcrNewItemNumberSetLotSize"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =660
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =960
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3240
            BackColor =14211288
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2160
                    Top =180
                    Width =2880
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="OracleItemNumber"
                    ControlSource ="OracleItemNumber"
                    AfterUpdate ="mcrNewItemNumberSetLotSize"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =300
                    Top =180
                    Width =1800
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label44"
                    Caption ="Oracle Item Number"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2160
                    Top =540
                    Width =2880
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="CatalogNumber"
                    ControlSource ="CatalogNumber"
                    AfterUpdate ="mcrNewItemNumberSetLotSize"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =300
                    Top =540
                    Width =1800
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label46"
                    Caption ="Catalog Number"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2160
                    Top =900
                    Width =2880
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Customer"
                    ControlSource ="Customer"
                    AfterUpdate ="mcrNewItemNumberSetLotSize"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =300
                    Top =900
                    Width =1800
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label48"
                    Caption ="Customer"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =2160
                    Top =1320
                    Width =240
                    TabIndex =3
                    Name ="Active"
                    ControlSource ="Active"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =300
                    Top =1260
                    Width =1800
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label51"
                    Caption ="Active"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1500
                    Top =2700
                    Width =2640
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="cmdCopyItem"
                    Caption ="Copy Item/BOM/Routing"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =2700
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =3060
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =14400
                    Left =2040
                    Top =2280
                    Width =2880
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ColumnInfo ="\"Oracle Item Number\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboItemNumberCopy"
                    RowSourceType ="Table/Query"
                    RowSource ="qrySelectItemNumber"
                    ColumnWidths ="2160;6480"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =300
                    Top =2280
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label14"
                    Caption ="Select Item to Copy"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =2160
                    Top =1680
                    Width =240
                    TabIndex =6
                    Name ="ExistingItem"
                    ControlSource ="ExistingItem"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =300
                    Top =1620
                    Width =1800
                    Height =300
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label53"
                    Caption ="Existing Item"
                    FontName ="Times New Roman"
                    ForeThemeColorIndex =0
                End
            End
        End
        Begin FormFooter
            Height =660
            BackColor =14211288
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2280
                    Top =60
                    Width =1020
                    FontSize =10
                    FontWeight =700
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

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

Private Sub SelectItemCopy()


    Dim strExt As String
    
    strExt = Mid(NewItemNumber, 8, 5)
    
    
    If strExt Like "*bx*" Then
        Me.cboItemNumberCopy.rowSource = "qrySelectItemNumberBox"
    
    ElseIf strExt Like "*ia*" Then
        Me.cboItemNumberCopy.rowSource = "qrySelectItemNumberArmor"
        
        'ElseIf strExt Like "*test*" Then
        'Me.cboItemNumberCopy.RowSource = "qrySelectItemNumberTEST"
    
    Else
        Me.cboItemNumberCopy.rowSource = "qrySelectItemNumber"
    
    End If
    



End Sub

Public Sub cboItemNumberCopy_AfterUpdate()

    [Forms]![frmMainMenu]![ItemNumberSelect] = [Forms]![frmNewItemNumber]![cboItemNumberCopy]

End Sub

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.name
    Forms!frmMainMenu.Visible = True
    DoCmd.Close acForm, "frmMainMenu"
End Sub

Private Sub cmdCopyItem_Click()

    If IsNull(Me.cboItemNumberCopy) = True Then
        MsgBox "Select an item to copy", vbCritical
    End If


    Call NewItemNumberContinue

End Sub

Private Sub Form_Open(Cancel As Integer)



    DoCmd.SetWarnings False
    Me.move 0, 0

    [Forms]![frmMainMenu]![CopyCount] = 0
    [Forms]![frmMainMenu]![CopyItem] = 0
    [Forms]![frmMainMenu]!ItemNumberSelect = Null

    'CreateSimpleShortcutMenu (Me.Name)
End Sub

Public Sub NewItemNumber_AfterUpdate()

    On Error GoTo Error:

    [Forms]![frmMainMenu]![NewItemNumber] = Me.NewItemNumber
    Me.Requery

    DoCmd.OpenForm "frmPremiseLotSize"
    Me.LotSize = Forms![frmPremiseLotSize]![LotSize]
    Forms![frmMainMenu]![StdLotSizeNew] = Me.LotSize
    DoCmd.Close acForm, "frmPremiseLotSize"


    Call NewItemNumberSetNewItem
    Call ItemType
    Call SelectItemCopy

ErrorExit:
    Exit Sub

Error:
    MsgBox Err.Description & " " & Err.Number
    Resume ErrorExit


End Sub

Private Sub ItemType()

    Dim NewItemType As String

    Select Case True

    Case Me.NewItemNumber Like "PT*"



        NewItemType = "Premise TB"

    Case Me.NewItemNumber Like "rbn*osp"

        NewItemType = "Ribbon OSP"

    Case Me.NewItemNumber Like "RBN*"
        NewItemType = "Premise Ribbon"


    Case Me.NewItemNumber Like "PE*"

        If Me.CatalogNumber Like "EM004?09???Q*" Then
            NewItemType = "A-BT"
        Else
            NewItemType = "Premise Sub"
        End If
    
    Case Me.NewItemNumber Like "DNT*"            ' Or Me.NewItemNumber Like "WTC*"
        NewItemType = "WTC FG"

    Case Me.NewItemNumber Like "WTC*"            ' Or Me.NewItemNumber Like "WTC*"
        NewItemType = "WTC SUB"
    
    Case Else
        NewItemType = "Premise FG"

    End Select

    Me.NewItemType = NewItemType
    Forms!frmMainMenu!NewItemType = NewItemType

    Me.cboItemNumberCopy = Null
    Me.cboItemNumberCopy.Requery

    If Me.NewItemNumber Like "*ia" Then
        Forms!frmMainMenu!StdLotSizeNew = "2200"
        Me.LotSize = "2200"
    ElseIf Me.NewItemNumber Like "dnt*" Or Me.NewItemNumber Like "wtc*" Then
        Me.LotSize = "12800"
    End If

    If Me.NewItemNumber Like "rbn*" Or Me.NewItemNumber Like "p*bx*" Then
        Forms!frmMainMenu!StdLotSizeNew = "1"
        Me.LotSize = "1"
    End If

End Sub

Private Sub NewItemType_AfterUpdate()
    Forms!frmMainMenu!NewItemType = Me.NewItemType

End Sub

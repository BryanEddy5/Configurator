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
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =735
    Top =3075
    Right =7860
    Bottom =9870
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa9dd3978168be440
    End
    RecordSource ="qryItemCategoryInterface"
    Caption ="frmItemInterface"
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
            Height =720
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =120
                    Top =420
                    Width =480
                    Height =300
                    FontSize =10
                    Name ="Label14"
                    Caption ="Org"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =660
                    Top =420
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
                    Left =2220
                    Top =420
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label16"
                    Caption ="Description"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =3600
                    Top =420
                    Width =2820
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Value"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =120
                    Top =120
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Item Categories"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =5820
                    Top =60
                    Width =1260
                    Height =600
                    FontSize =10
                    FontWeight =700
                    ForeColor =-2147483630
                    Name ="CmdEnterNewCat"
                    Caption ="New Product Category"
                    OnClick ="mcrOpenNewProductCat"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =4500
                    Top =60
                    Width =1260
                    Height =600
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =-2147483630
                    Name ="cmdBlanks"
                    Caption ="Remove\015\012Blanks"
                    OnClick ="mcrCategoryBlanksDelete"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2220
                    Width =1005
                    Height =405
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="Command29"
                    Caption ="Edit Cat"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =366
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =120
                    Top =15
                    Width =480
                    Height =300
                    ColumnWidth =465
                    ColumnOrder =3
                    FontSize =10
                    Name ="OrganizationCode"
                    ControlSource ="OrganizationCode"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =660
                    Top =15
                    Width =1500
                    Height =300
                    ColumnWidth =1185
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =1
                    Name ="ItemNumber"
                    ControlSource ="ItemNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =2220
                    Top =15
                    Width =1320
                    Height =300
                    ColumnWidth =3000
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =2
                    Name ="CategorySetName"
                    ControlSource ="CategorySetName"
                    FontName ="Times New Roman"

                End
                Begin CheckBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =6480
                    Top =75
                    ColumnOrder =4
                    TabIndex =4
                    Name ="AllowOverride"
                    ControlSource ="AllowOverride"

                End
                Begin ComboBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =8640
                    Left =3600
                    Top =15
                    Width =2760
                    Height =300
                    ColumnWidth =465
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="CategoryName"
                    ControlSource ="CategoryName"
                    RowSourceType ="Table/Query"
                    RowSource ="qryItemInterfaceCategoryValues"
                    ColumnWidths ="2880;5760"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =119
                    Left =6720
                    Top =30
                    Width =366
                    Height =306
                    TabIndex =5
                    Name ="cmdDelete"
                    Caption ="Command25"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
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

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
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

Private Sub CategoryName_AfterUpdate()
    Me.CategoryName.Requery
End Sub

Private Sub CategoryName_GotFocus()
    Me.CategoryName.Requery
End Sub

Private Sub cmdDelete_Click()

    If Me.CategorySetName <> "Product Class" Then
        DeleteARecord
    End If

    

End Sub

Private Sub Command29_Click()
    On Error GoTo Err_Command29_Click

    Dim stDocName As String

    stDocName = "qryItemCategoryInterface"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_Command29_Click:
    Exit Sub

Err_Command29_Click:
    MsgBox Err.Description
    Resume Exit_Command29_Click
    
End Sub

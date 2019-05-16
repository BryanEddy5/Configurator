Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =23
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5640
    DatasheetFontHeight =10
    Left =9450
    Top =1920
    Right =17415
    Bottom =11805
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd42c1f838f90e440
    End
    RecordSource ="tblFiberConversion_CopyingOldItemsToNew"
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
            Height =1140
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =900
                    Top =780
                    Width =1560
                    Height =300
                    FontSize =10
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label4"
                    Caption ="New Fiber Number"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2760
                    Top =180
                    Width =1380
                    Height =420
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="cmdClose"
                    Caption ="Close"
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
                    Left =1200
                    Top =180
                    Width =1380
                    Height =420
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdClearItems"
                    Caption ="Clear Items"
                    OnClick ="mcrItemToCopyClearAll"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =2760
                    Top =780
                    Width =1860
                    Height =300
                    FontSize =10
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label28"
                    Caption ="Existing Fiber Number"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =240
                    Top =240
                    Width =300
                    Height =300
                    FontSize =10
                    TabIndex =2
                    BackColor =255
                    ForeColor =-2147483640
                    Name ="MasterLevelTemplateDefault"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            Height =306
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4740
                    Width =306
                    Height =306
                    ForeColor =0
                    Name ="cmdDelete"
                    Caption ="Command16"
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

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =900
                    Width =1800
                    Height =300
                    ColumnWidth =1605
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="NewItemNumber"
                    ControlSource ="New BaseItem"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2760
                    Width =1800
                    Height =300
                    ColumnWidth =1935
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="ExistingItemNumber"
                    ControlSource ="Old BaseItem"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            Height =120
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

Private Sub cmdClose_Click()


    DoCmd.Close acForm, "frmFiberConvItems"
    
End Sub

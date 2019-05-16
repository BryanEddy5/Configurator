Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =85
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =21186
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =1140
    Top =2790
    Right =22980
    Bottom =5625
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc54d20bfb4dce440
    End
    RecordSource ="tblNewBasicCableHold_New"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
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
        Begin FormHeader
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =1740
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Unit Color Series"
                    Caption ="New Oracle Part #"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Unit_Color_Series"
                End
                Begin Label
                    OverlapFlags =85
                    Left =4440
                    Top =60
                    Width =4200
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Unit ID"
                    Caption ="Item No"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Unit_ID"
                    LayoutCachedLeft =4440
                    LayoutCachedTop =60
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =85
                    Left =8760
                    Top =60
                    Width =1800
                    Height =240
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Description"
                    Caption ="Base"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =8760
                    LayoutCachedTop =60
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =85
                    Left =10680
                    Top =60
                    Width =2040
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Unit Type"
                    Caption ="Jacket Color"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Unit_Type"
                    LayoutCachedLeft =10680
                    LayoutCachedTop =60
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =85
                    Left =12840
                    Top =60
                    Width =2220
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Print type"
                    Caption ="Fiber Type"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Print_type"
                    LayoutCachedLeft =12840
                    LayoutCachedTop =60
                    LayoutCachedWidth =15060
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =85
                    Left =2040
                    Top =60
                    Width =2220
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Label0"
                    Caption ="Customer"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =85
                    Left =15180
                    Top =60
                    Width =2220
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Label8"
                    Caption ="Fiber Type 2"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =15180
                    LayoutCachedTop =60
                    LayoutCachedWidth =17400
                    LayoutCachedHeight =315
                End
            End
        End
        Begin Section
            Height =426
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =120
                    Width =1740
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txOracleItem"
                    ControlSource ="New Oracle Part #"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4440
                    Top =120
                    Width =4200
                    Height =255
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text1"
                    ControlSource ="Item No"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =120
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8760
                    Top =120
                    Width =1800
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text2"
                    ControlSource ="Base"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =120
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10680
                    Top =120
                    Width =2040
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text4"
                    ControlSource ="Jacket Color"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =10680
                    LayoutCachedTop =120
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12840
                    Top =120
                    Width =2280
                    Height =255
                    FontWeight =700
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text5"
                    ControlSource ="Fiber Oracle item"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =12840
                    LayoutCachedTop =120
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2040
                    Top =120
                    Width =2220
                    Height =255
                    FontWeight =700
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text6"
                    ControlSource ="Customer"
                    FontName ="MS Sans Serif"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =20760
                    Top =60
                    Width =366
                    Height =366
                    TabIndex =6
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

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15180
                    Top =120
                    Width =2280
                    Height =255
                    FontWeight =700
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text9"
                    ControlSource ="FiberType2"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =15180
                    LayoutCachedTop =120
                    LayoutCachedWidth =17460
                    LayoutCachedHeight =375
                End
            End
        End
        Begin FormFooter
            Height =480
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =240
                    Width =1380
                    Height =480
                    FontWeight =700
                    ForeColor =0
                    Name ="btnViewFam"
                    Caption ="View Faimly"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =1740
                    Width =1380
                    Height =480
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="btnLoadCable"
                    Caption ="Load"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

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

Private Sub btnLoadCable_Click()
    'DoCmd.OpenQuery "qryLOADPremisCableSheath"
    'DoCmd.OpenQuery "qryBasicCableSheethActivate"
    Me.Refresh

    DoCmd.OpenQuery "qryLOADBasicCable"

    Me.btnViewFam.SetFocus
    Me.btnLoadCable.Enabled = False


End Sub

Private Sub Command32_Click()
    DoCmd.OpenQuery "qryBasicSheathFamilyFind", , acReadOnly
End Sub

Private Sub btnViewFam_Click()
    DoCmd.OpenQuery "qryBasicSheathFamilyFind", , acReadOnly

End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Err_cmdDelete_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
End Sub

Private Sub Form_Load()
    Me.AllowAdditions = True
    DoCmd.GoToRecord , , acNewRec
    Me.AllowAdditions = False
    Me.btnLoadCable.Enabled = True
    
End Sub

Private Sub txOracleItem_AfterUpdate()

    'If the Item Request Item Number has changed then the item being inserted into the creation log needs to be updated
    If CreateNewItem.NewItemNumber = Me.txOracleItem.OldValue Then

        CreateNewItem.NewItemNumber = Me.txOracleItem.value

    End If



End Sub

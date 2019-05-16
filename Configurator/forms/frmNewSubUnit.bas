Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =7
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19806
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =3360
    Top =930
    Right =22980
    Bottom =4215
    DatasheetGridlinesColor =12632256
    OrderBy ="[tblNewProductConstructionHold].[New Oracle Part #]"
    RecSrcDt = Begin
        0x4a9379be7691e440
    End
    RecordSource ="tblNewProductConstructionHold"
    OnClose ="[Event Procedure]"
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
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =480
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =120
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
                    Left =2100
                    Top =120
                    Width =2880
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Unit ID"
                    Caption ="Item No"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Unit_ID"
                End
                Begin Label
                    OverlapFlags =85
                    Left =5040
                    Top =120
                    Width =1800
                    Height =240
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Description"
                    Caption ="Base"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =85
                    Left =6960
                    Top =120
                    Width =2580
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="TB Color Series"
                    Caption ="Jacket Material"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="TB_Color_Series"
                End
                Begin Label
                    OverlapFlags =85
                    Left =9660
                    Top =120
                    Width =2040
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Unit Type"
                    Caption ="Jacket Color"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Unit_Type"
                    LayoutCachedLeft =9660
                    LayoutCachedTop =120
                    LayoutCachedWidth =11700
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =11820
                    Top =120
                    Width =2220
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Print type"
                    Caption ="Fiber Type"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Print_type"
                    LayoutCachedLeft =11820
                    LayoutCachedTop =120
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =14280
                    Top =120
                    Width =2220
                    Height =255
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Label13"
                    Caption ="Fiber Type 2"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =14280
                    LayoutCachedTop =120
                    LayoutCachedWidth =16500
                    LayoutCachedHeight =375
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =366
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =60
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
                    Left =2040
                    Top =60
                    Width =2880
                    Height =255
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text1"
                    ControlSource ="Item No"
                    FontName ="MS Sans Serif"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4980
                    Top =60
                    Width =1800
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text2"
                    ControlSource ="Base"
                    FontName ="MS Sans Serif"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6900
                    Top =60
                    Width =2580
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text3"
                    ControlSource ="Jacket Material"
                    FontName ="MS Sans Serif"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9600
                    Top =60
                    Width =2040
                    Height =255
                    FontWeight =700
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text4"
                    ControlSource ="Jacket Color"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =60
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11760
                    Top =60
                    Width =2280
                    Height =255
                    ColumnWidth =3870
                    FontWeight =700
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text5"
                    ControlSource ="Fiber Oracle item"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =11760
                    LayoutCachedTop =60
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =315
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16800
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
                    ControlTipText ="Delete Record"

                    LayoutCachedLeft =16800
                    LayoutCachedWidth =17166
                    LayoutCachedHeight =366
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14220
                    Top =60
                    Width =2280
                    Height =255
                    FontWeight =700
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Text14"
                    ControlSource ="FiberType2"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =14220
                    LayoutCachedTop =60
                    LayoutCachedWidth =16500
                    LayoutCachedHeight =315
                End
            End
        End
        Begin FormFooter
            Height =660
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Height =480
                    FontWeight =700
                    Name ="btnViewFam"
                    Caption ="View Family"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =17400
                    Top =120
                    Height =480
                    FontWeight =700
                    TabIndex =1
                    Name ="Command11"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Default = NotDefault
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =1740
                    Top =120
                    Width =1380
                    Height =480
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="Command32"
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

Private Sub btnViewFam_Click()
    DoCmd.OpenQuery "qryNewBasicProductFindFamily", , acReadOnly
End Sub

Private Sub Command11_Click()
    On Error GoTo Err_cmdCloseForm_Click

    DoCmd.RunCommand acCmdSaveRecord
    
    If CurrentProject.AllForms("frmFinishNewSubunit").IsLoaded Then
        DoCmd.Close acForm, "frmFinishNewSubunit"
    Else
        DoCmd.Close acForm, "frmNewSubUnit"
    
    End If
                
    Form_Close
Exit_cmdCloseForm_Click:
    Exit Sub

Err_cmdCloseForm_Click:
    MsgBox Err.Description
    Resume Exit_cmdCloseForm_Click
End Sub

Private Sub Command32_Click()

    Me.Refresh
    DoCmd.OpenQuery "qryLOADBasicProductConstruction"
    'DoCmd.OpenQuery "qryLOADPremisBasicConstruction"
    'DoCmd.OpenQuery "qryBasicCableConstActivate"
    Me.Command11.SetFocus
    Me.Command32.Enabled = False

    Call CreateNewItem.InsertRequestToLog



    Beep
    MsgBox "Load Complete"

    Call Command11_Click

End Sub

Private Sub Form_Close()
    Dim formToClose$

    formToClose = "frmCreateItem"

    If CurrentProject.AllForms(formToClose).IsLoaded Then Call Forms(formToClose).btnClose_Click

    'DoCmd.Close acForm, Me.Name

End Sub

Private Sub Form_Load()
    DoCmd.Maximize
    'Me.AllowAdditions = True
    'DoCmd.GoToRecord , , acNewRec
    Me.AllowAdditions = False
    'Me.btnViewFam.SetFocus
    Me.Command32.Enabled = True
    
    If CurrentProject.AllForms("frmFinishNewSubunit").IsLoaded Then
        Me.Command11.Visible = False
        Me.Command32.Visible = False
    Else
        Me.Command11.Visible = True
        Me.Command32.Visible = True
    End If
    

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

Private Sub txOracleItem_AfterUpdate()

    'If the Item Request Item Number has changed then the item being inserted into the creation log needs to be updated
    If CreateNewItem.NewItemNumber = Me.txOracleItem.OldValue Then

        CreateNewItem.NewItemNumber = Me.txOracleItem.value

    End If

End Sub

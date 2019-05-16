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
    Width =7620
    DatasheetFontHeight =10
    ItemSuffix =85
    Left =540
    Top =885
    Right =9210
    Bottom =6690
    DatasheetGridlinesColor =12632256
    Filter ="([Standards_CableConstructionStandards].[Standard] Like \"*tpr*\")"
    RecSrcDt = Begin
        0x59fd25514c19e540
    End
    RecordSource ="Standards_CableConstructionStandards"
    Caption ="Item Search"
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
                    Left =240
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="BaseID"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =1680
                    Width =3060
                    Name ="Box35"
                    LayoutCachedLeft =1680
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =720
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1260
                    Width =7020
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =1800
                    Top =120
                    Width =2760
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Cable Standards"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =1800
                    LayoutCachedTop =120
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =615
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1140
                    Top =840
                    Width =1320
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =840
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1200
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2760
                    Top =840
                    Width =1320
                    TabIndex =1
                    Name ="btnAdd"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2760
                    LayoutCachedTop =840
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1200
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =840
                    Width =1320
                    TabIndex =2
                    Name ="btnSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4380
                    LayoutCachedTop =840
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1200
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =1920
                    Top =1680
                    Width =4680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label80"
                    Caption ="Standard"
                    LayoutCachedLeft =1920
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =1980
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5340
                    Top =360
                    Height =255
                    TabIndex =3
                    Name ="BaseID"

                    LayoutCachedLeft =5340
                    LayoutCachedTop =360
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =615
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
                    Left =240
                    Top =60
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txBaseID"
                    ControlSource ="BaseID"

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7020
                    Top =60
                    Width =480
                    Height =300
                    TabIndex =2
                    Name ="btnDelete"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddddddddddddddddddddddddddddddd0fd ,
                        0xddd0fddddddddddddd000fdddddd0fdddd000fddddd0fdddddd000fddd00fddd ,
                        0xdddd000fd00fddddddddd00000fddddddddddd000fddddddddddd00000fddddd ,
                        0xdddd000fd0fddddddd0000fddd00fdddd0000fddddd00fddd00fdddddddd00fd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
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

                    LayoutCachedLeft =7020
                    LayoutCachedTop =60
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =4680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbStandard"
                    ControlSource ="Standard"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Standards_IndustryStandards.Standard FROM Standards_IndustryStandards;"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =6600
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
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1920
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="Refresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

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
Private ID As Integer

Private Sub btnAdd_Click()
Me.AllowAdditions = True
Call btnEdit_Click
ID = Nz(Me.txBaseID, Me.BaseID)
End Sub

Private Sub btnClose_Click()
DoCmd.Close
End Sub

Private Sub btnDelete_Click()
Call DeleteARecord
End Sub

Private Sub btnEdit_Click()
Dim ctl As control
For Each ctl In Me.Controls
    If ctl.ControlType = acComboBox Then
        If ctl.Locked <> False Then
            ctl.Locked = False
        End If
    End If
Next ctl
Me.txBaseID.Locked = True
ID = Nz(Me.txBaseID, Me.BaseID)
End Sub


Private Sub btnSave_Click()
Me.AllowAdditions = False
Me.Requery
End Sub

Private Sub cbStandard_AfterUpdate()
Me.txBaseID = ID
End Sub

Private Sub Form_Open(Cancel As Integer)
Me.AllowAdditions = False
End Sub

Private Sub Refresh_Click()
Me.Requery
Me.Refresh
End Sub

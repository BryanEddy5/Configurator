Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4200
    DatasheetFontHeight =10
    ItemSuffix =73
    Left =5235
    Top =1140
    Right =9435
    Bottom =7065
    DatasheetGridlinesColor =12632256
    Filter ="Tab = '9723907' AND Section = '9' AND Construction = '-'"
    RecSrcDt = Begin
        0x2f2c92f912abe440
    End
    RecordSource ="tblBurnListingType"
    Caption ="Burn Listing"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
            Height =2040
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =300
                    Top =180
                    Width =2430
                    Height =780
                    FontSize =24
                    FontWeight =500
                    Name ="Label63"
                    Caption ="Burn Print"
                    FontName ="Times New Roman"
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =93
                    Left =300
                    Top =1680
                    Width =1500
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label70"
                    Caption ="Burn Print"
                End
                Begin Line
                    OverlapFlags =87
                    Top =1980
                    Width =4140
                    Name ="Line72"
                End
            End
        End
        Begin Section
            Height =486
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =1500
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="Text71"
                    ControlSource ="Type"

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

Private Sub btnFiberClose_Click()
    DoCmd.Close acForm, "frmFindItemSubunits"
End Sub

Private Sub btnFind_Click()

End Sub

Private Sub btnClose_Click()
    DoCmd.Close acForm, "frmFindBurnListing_Print"
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Error:





Error_Exit:
    Exit Sub
Error:
    MsgBox ("No Burn Listing"), vbCritical
    Resume Error_Exit

End Sub

Private Sub Refresh_Click()
    Me.Requery
    Me.Refresh
End Sub

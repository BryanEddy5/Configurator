Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5760
    DatasheetFontHeight =10
    ItemSuffix =76
    Left =5100
    Top =885
    Right =11115
    Bottom =4110
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8364285936d8e440
    End
    Caption ="Email"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    Moveable =0
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =3
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =90.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =15790320
            BackThemeColorIndex =6
            BackTint =30.0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
        End
        Begin FormHeader
            Height =960
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    OverlapFlags =93
                    Left =240
                    Top =60
                    Width =2460
                    Name ="Box35"
                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =780
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =840
                    Width =5160
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =840
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =360
                    Top =180
                    Width =2220
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblApproval"
                    Caption ="Add Email"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =360
                    LayoutCachedTop =180
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =675
                End
            End
        End
        Begin Section
            Height =780
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =420
                    Width =3180
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="txEmailAddress"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =420
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="txUserName"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =360
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="User Name"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =360
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =420
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Email Address"
                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =720
                End
            End
        End
        Begin FormFooter
            Height =1500
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =660
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =720
                    Width =1320
                    TabIndex =1
                    Name ="btnSubmit"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3960
                    LayoutCachedTop =720
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =1080
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =600
                    Top =180
                    Width =4020
                    Height =405
                    FontSize =9
                    FontWeight =700
                    ForeColor =2366701
                    Name ="lblEmailNote"
                    Caption ="Email sent.  Please wait for email confirmation."
                    FontName ="Times New Roman"
                    LayoutCachedLeft =600
                    LayoutCachedTop =180
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =585
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


Private Sub btnSubmit_Click()
Dim start As Double
Dim Admin As New clsUser
Dim adminEmail$


adminEmail = Admin.getAdminEmailsSemiColonDelimited

DoCmd.SetWarnings False
DoCmd.Hourglass True



On Error GoTo ErrorHandler

If IsNull(Me.txEmailAddress) Or Not Me.txEmailAddress Like "*@aflglobal.com*" Then

    MsgBox "Email address is a required field.  Ensure you are entering an AFL Global email address.", vbCritical
    DoCmd.Hourglass False
    Exit Sub
End If

Call AddEmailAddress(sUserName, Me.txEmailAddress)



Call SendEmail("Email Add", Me.txEmailAddress, "Your email address has been successfully added to the Premise Program.", , adminEmail)

lblEmailNote.Visible = True

start = Timer
While Timer < start + 2
    DoEvents
Wend


If MsgBox("Did you receive an email notification?", vbYesNo + vbCritical) = vbNo Then
    lblEmailNote.Visible = False
    MsgBox "Please re-enter your email address"
    DoCmd.Hourglass False
    Exit Sub
End If



DoCmd.Hourglass False
MsgBox "Success!"

DoCmd.Close acForm, Me.name

Call Core_Login.OpenRoleBasedForm



ErrorExit:
Exit Sub


ErrorHandler:
MsgBox Err.Number & " " & Err.Description
DoCmd.Hourglass False
Resume ErrorExit

End Sub

Private Sub Form_Load()
lblEmailNote.Visible = False
Me.txUserName = sUserName

End Sub

Private Sub txEmailAddress_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Me.btnSubmit.SetFocus
    Call btnSubmit_Click
End If



End Sub

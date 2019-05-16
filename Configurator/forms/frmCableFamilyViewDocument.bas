Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    TabularFamily =220
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =10785
    Top =3495
    Right =17715
    Bottom =7995
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7732119c16abe440
    End
    Caption ="Revision Reason"
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
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =3540
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =1905
                    Top =120
                    Width =3465
                    Height =480
                    FontSize =14
                    Name ="Label2"
                    Caption ="Choose Doc"
                    LayoutCachedLeft =1905
                    LayoutCachedTop =120
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1920
                    Top =2280
                    Width =1200
                    TabIndex =1
                    Name ="btnOpen"
                    Caption ="Open"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =2280
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =2640
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4020
                    Top =2280
                    Width =1200
                    TabIndex =2
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =2280
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2640
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Top =960
                    Width =4380
                    Height =1260
                    FontSize =10
                    Name ="bxDocList"
                    RowSourceType ="Value List"
                    RowSource ="1774.U.pdf;1774.V.pdf;1774.X.pdf;1774.Y.pdf;1774.Z.pdf"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =1380
                    LayoutCachedTop =960
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =2220
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

Private Sub btnClose_Click()
DoCmd.Close
End Sub

Private Sub btnOpen_Click()
Dim CustomerFile As String
Dim path As String

On Error GoTo ErrorHandler:

CustomerFile = Me.bxDocList
path = Core_PathwayManager.GetPathway(NPR)
path = path & Me.bxDocList

Call OpenCustomerSpec(path, Me.bxDocList)



ErrorExit:
Exit Sub

ErrorHandler:
If Err.Number = 94 Then
    MsgBox "Please choose a document"
Else
    MsgBox Err.Description & " " & Err.Number
End If
Resume ErrorExit



End Sub

Private Sub bxDocList_DblClick(Cancel As Integer)
Call btnOpen_Click
End Sub

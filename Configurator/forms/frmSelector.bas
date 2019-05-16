Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
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
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5040
    ItemSuffix =8
    Left =8790
    Top =2190
    Right =16410
    Bottom =10905
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xc07df631d786e340
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnTimer ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =10
            BorderColor =-2147483633
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =2
        End
        Begin Image
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
            SpecialEffect =2
            BackStyle =1
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BackColor =-2147483633
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =10
            FontName ="Tahoma"
            AsianLineBreak =1
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
            FontSize =10
            FontName ="Tahoma"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
            FontSize =10
            FontName ="Tahoma"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            BackStyle =0
            OldBorderStyle =1
            BackColor =-2147483633
        End
        Begin CustomControl
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Tab
            FontSize =10
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            Height =5520
            BackColor =-2147483633
            Name ="Detail"
            AlternateBackColor =-2147483633
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =120
                    Width =3240
                    Height =780
                    FontSize =8
                    TabIndex =4
                    Name ="tbPrompt"

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =900
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =3480
                    Top =120
                    TabIndex =2
                    Name ="cmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =120
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3480
                    Top =540
                    TabIndex =3
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =540
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =900
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =1020
                    Width =4800
                    Height =285
                    FontSize =8
                    LeftMargin =21
                    TopMargin =21
                    RightMargin =21
                    ForeColor =9211020
                    Name ="tbSearchText"
                    Format ="@;\" <enter search text>\""
                    OnChange ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =1305
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffffff00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =120
                    Top =1440
                    Width =4795
                    Height =3595
                    AutoActivate =1
                    TabIndex =1
                    Name ="ListView0"
                    OleData = Begin
                        0x000e0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffdfffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000feffffff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000052006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000004bf0d1bd8b85d111b16a00c0 ,
                        0xf0283628000000000000000000000000d05bf3edb55fcc010500000040010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000200000093000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000000feffffff0300000004000000feffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000039333638 ,
                        0x323635452d383546452d313164312d384245332d303030304638373534444131 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000004bf0d1bd8b85d111b16a00c0f02836282143341208000000c3200000 ,
                        0x7f1800004e087deb010006001c000000000000000000000000860100c3200000 ,
                        0x01efcdab00000500883f220e0280ffff0800008005000080b0e1060e00000000 ,
                        0x00000000000000001fdeecbd0100050029e1060e0352e30b918fce119de300aa ,
                        0x004bb85101000000900144420100065461686f6d610000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="ListViewCtrl"
                    Class ="MSComctlLib.ListViewCtrl.2"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4915
                    LayoutCachedHeight =5035
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =5220
                    Width =4380
                    Height =240
                    FontSize =8
                    ForeColor =4144959
                    Name ="lbMessage"
                    Caption ="| Type to search  | Enter to select  | Esc to cancel"
                    LayoutCachedLeft =120
                    LayoutCachedTop =5220
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =5460
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


'********************************************************************
'*                                                                                                                                   Oct 7 2009
'*       USAGE NOTE:  This form must be opened by Function Show() of the cSelector class
'*
'********************************************************************
Public Event Selected(value As String)

Private WithEvents m_lvw As MSComctlLib.Listview
Attribute m_lvw.VB_VarHelpID = -1
Private m_lastKeyTime As Single
Private WithEvents m_builder As cListViewBuilder
Attribute m_builder.VB_VarHelpID = -1

'********************************************************************
'*                                                                                                   Jul 6 2005
'*       Properties
'*
'********************************************************************


Friend Property Get Listview() As MSComctlLib.Listview
   If m_lvw Is Nothing Then Set m_lvw = Me.ListView0.Object
   Set Listview = m_lvw
End Property

Friend Property Get value() As Variant
   If Listview.SelectedItem Is Nothing Then
      value = Null
   Else
      value = Listview.SelectedItem.Text
   End If
End Property
Friend Property Get ListViewBuilder() As cListViewBuilder
    If m_builder Is Nothing Then Set m_builder = New cListViewBuilder
    Set ListViewBuilder = m_builder
End Property

Friend Property Get ValueArray() As Variant
   Dim var
   If Me.Listview.SelectedItem Is Nothing Then
      ReDim var(0)
      var(0) = Null
   Else
      Dim i As Integer
      Dim Count As Integer
      Count = m_lvw.ColumnHeaders.Count
      ReDim var(Count - 1)
      For i = 1 To Count
         If i = 1 Then
            var(0) = Me.Listview.SelectedItem.Text
         Else
            var(i - 1) = Me.Listview.SelectedItem.ListSubItems(i - 1).Text
         End If
      Next
   End If
   ValueArray = var
End Property


'********************************************************************
'*                                                                                                   Dec 14 2004
'*       Form Operations
'*
'********************************************************************
Private Sub Form_Open(Cancel As Integer)
'On Error GoTo handler
   
   'set a reference to Me in the selectors
   'resize and center Me
   DoMoveResize

    

      'set various display properties of the form
      Me.Caption = "Item Selector"
      Me.tbPrompt = "Please select an Item from the list.  Notice that as you type in search text field, the list is filtered."
      ListViewBuilder.Load(Me.ListView0, "SELECT Oracle, ItemNo FROM vfindCutSheet", "1;3;3;3;2").Rebuild
      'if the FindItem option exists, try to find it
      Call FindListItem

   
   'set form level reference to the the list view object in the control
   Set m_lvw = Me.ListView0.Object
   '... and make sure nothing is selected
   Set m_lvw.SelectedItem = Nothing
   
'final:
'   Screen.MousePointer = 0
'   Exit Sub
'
'handler:
'   Screen.MousePointer = 0
'   Me.Visible = False
'   Err.Raise Err.Number, Err.Source & " fSelector.Form_Open()", Err.Description
End Sub


'********************************************************************
'*                                                                                                   Apr 23 2005
'*       Command Button Event Handlers
'*
'********************************************************************
Private Sub cmdCancel_Click()
   Set Listview.SelectedItem = Nothing
   Me.Visible = False
   DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdOK_Click()
   If Listview.SelectedItem Is Nothing Then
      MsgBox "Please click list item to select it", vbInformation, "Nothing Selected"
   Else
      RaiseEvent Selected(Nz(Me.value, 0))
      Me.Visible = False
      DoCmd.Close acForm, Me.name
   End If
   
   
End Sub


'********************************************************************
'*                                                                                                                                   Jan 21 2010
'*       List View Event Handling   >
'*
'********************************************************************
Private Sub m_lvw_DblClick()
'  Dbl_Click selects item and closes form
   cmdOK_Click
End Sub

'********************************************************************
'*                                                                                                   Apr 23 2005
'*       Search Text Change
'*
'********************************************************************
Private Sub tbSearchText_Change()
   'we do a little trick here to make sure this works quickly,
   'set timer interval for every keystroke
   Me.TimerInterval = 1
End Sub

Private Sub Form_Timer()
   Dim where As String
   Me.TimerInterval = 0
   If Me.tbSearchText.Text <> "" Then where = GetWhereClause(Me.tbSearchText.Text)
   ListViewBuilder.Rebuild where
End Sub

'********************************************************************
'*                                                                                                                                   Aug 20 2011
'*       Utilities>
'*
'********************************************************************
Private Sub DoMoveResize()
   Const BORDER_WIDTH = 120
   Dim inchWidth#
   Dim inchHeight#
   inchWidth = 5.5 * 1440
   inchHeight = 6.5 * 1440

   'the form is small in design view and ends up getting resized based on calling code's settings,
   'so here we centered it on the screen
   Dim move(1) As Long
   With Me.Form
      'calculates how much difference there is between the design view size and the requested size, and divides by 2
      move(0) = Abs(inchWidth - Me.InsideWidth) / 2
      move(1) = Abs(inchHeight - Me.InsideHeight) / 4
      'then the form is moved and sized
      Me.move Me.WindowLeft - move(0), Me.WindowTop - move(1), inchWidth, inchHeight
   End With

   'set the height and width of the list box
   With Me.ListView0
      .Width = Me.InsideWidth - .Left - BORDER_WIDTH
      .Height = Me.InsideHeight - .top - BORDER_WIDTH * 3
      'make the tbSearchText textbox the same
      Me.tbSearchText.Width = .Width
   End With
   'move the form's buttons to the far right of the form
   Me.cmdOK.Left = Me.InsideWidth - Me.cmdOK.Width - BORDER_WIDTH
   Me.cmdCancel.Left = Me.cmdOK.Left
   'size the prompt
   Me.tbPrompt.Width = Me.InsideWidth - Me.tbPrompt.Left - Me.cmdOK.Width - BORDER_WIDTH * 2
   'move the label at the bottom
   Me.lbMessage.top = Me.InsideHeight - 0.2 * 1440

End Sub

Function GetWhereClause(Text As String) As String
'  Constructs a where clause to pass to the listview builder
   Dim i As Integer
   Dim where As String
   Dim prm As String
   
   prm = "'*" & Replace(Text, "'", "_") & "*'"
   For i = 1 To Listview.ColumnHeaders.Count
      If m_lvw.ColumnHeaders(i).Width > 0 Then
         where = where & "[" & m_lvw.ColumnHeaders(i).tag & "] LIKE " & prm & " Or "
      End If
   Next i
   'drop the trailing 'Or' if it exists
   If Right(where, 3) = "Or " Then where = Mid(where, 1, Len(where) - 3)
   GetWhereClause = where

End Function

Private Sub FindListItem()
'  If a FindItemID was specified in the selector call, this routine attempts to locate it.
   Dim item As MSComctlLib.ListItem
   Set item = Listview.FindItem(CStr(Me.Listview.SelectedItem), lvwText, , lvwWhole)
   If Not item Is Nothing Then
      item.Selected = True
      item.EnsureVisible
      On Error GoTo errIndexOutOfBounds
      Do While Not Listview.GetFirstVisible.Selected
         Set item = Listview.ListItems(item.Index + 1)
         item.EnsureVisible
      Loop
      Me.ListView0.SetFocus
      On Error GoTo 0
   End If
   Exit Sub
   
errIndexOutOfBounds:
   If Err = 35600 Then
      'ignore - index out of bounds
   Else
      Err.Raise Err, Err.Source & " - FindListItem()"
   End If

End Sub

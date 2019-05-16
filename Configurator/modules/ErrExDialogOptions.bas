Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *****************************************************************************
' *  VCOM library: vbWatchdog  http://www.everythingaccess.com/vbwatchdog.htm *
' *  3.8.2 (Ultimate Edition TRI built Feb 28 2018                            *
' *---------------------------------------------------------------------------*
' *  This is a demonstration version of the library for TESTING purposes only.*
' *                                                                           *
' *  This software is provided "as-is," without any express or implied        *
' *  warranty.  In no event shall the author be held liable for any           *
' *  consequential, incidental, direct, indirect, special, punitive, or other *
' *  damages arising from the use of this software (including without         *
' *  limitation, damages for loss of business profits, business interruption, *
' *  loss of information or other pecuniary loss).                            *
' *****************************************************************************

Option Explicit

Enum DefaultButtonActions
    BUTTONACTION_ONERRORRESUMENEXT = &H2
    BUTTONACTION_ONERROREND = &H4
    BUTTONACTION_ONERRORDEBUG = &H5
    BUTTONACTION_ONERRORRETRY = &H7
    BUTTONACTION_ONERROREXITPROCEDURE = &H9
    BUTTONACTION_SHOWVARIABLES = &HA
    BUTTONACTION_SHOWHELP = &HB
    BUTTONACTION_ONERRORUNWIND = &H11
    BUTTONACTION_ONERRORUNWINDNOFINALLY = &H12
    BUTTONACTION_VARIABLES_CLOSE = &H101
End Enum

Public VCOMObject As Object      ' DO NOT USE!!! THIS IS FOR INTERNAL USE ONLY!!!
Attribute VCOMObject.VB_VarUserMemId = -4

Property Get HTML_MainBody() As String
    HTML_MainBody = VCOMObject.HTML_MainBody
End Property

Property Let HTML_MainBody(ByVal value As String)
    VCOMObject.HTML_MainBody = value
End Property

Property Get HTML_MoreInfoBody() As String
    HTML_MoreInfoBody = VCOMObject.HTML_MoreInfoBody
End Property

Property Let HTML_MoreInfoBody(ByVal value As String)
    VCOMObject.HTML_MoreInfoBody = value
End Property

Property Get HTML_CallStackItem() As String
    HTML_CallStackItem = VCOMObject.HTML_CallStackItem
End Property

Property Let HTML_CallStackItem(ByVal value As String)
    VCOMObject.HTML_CallStackItem = value
End Property

Property Get HTML_VariableItem() As String
    HTML_VariableItem = VCOMObject.HTML_VariableItem
End Property

Property Let HTML_VariableItem(ByVal value As String)
    VCOMObject.HTML_VariableItem = value
End Property

Property Get WindowCaption() As String
    WindowCaption = VCOMObject.WindowCaption
End Property

Property Let WindowCaption(ByVal value As String)
    VCOMObject.WindowCaption = value
End Property

Property Get MinimumWindowWidth() As Long
    MinimumWindowWidth = VCOMObject.MinimumWindowWidth
End Property

Property Let MinimumWindowWidth(ByVal value As Long)
    VCOMObject.MinimumWindowWidth = value
End Property

Property Get MoreInfoCaption() As String
    MoreInfoCaption = VCOMObject.MoreInfoCaption
End Property

Property Let MoreInfoCaption(ByVal value As String)
    VCOMObject.MoreInfoCaption = value
End Property

Property Get LessInfoCaption() As String
    LessInfoCaption = VCOMObject.LessInfoCaption
End Property

Property Let LessInfoCaption(ByVal value As String)
    VCOMObject.LessInfoCaption = value
End Property

Property Get ButtonPaddingH() As Long
    ButtonPaddingH = VCOMObject.ButtonPaddingH
End Property

Property Let ButtonPaddingH(ByVal value As Long)
    VCOMObject.ButtonPaddingH = value
End Property

Property Get ButtonPaddingV() As Long
    ButtonPaddingV = VCOMObject.ButtonPaddingV
End Property

Property Let ButtonPaddingV(ByVal value As Long)
    VCOMObject.ButtonPaddingV = value
End Property

Property Get ButtonSpacingH() As Long
    ButtonSpacingH = VCOMObject.ButtonSpacingH
End Property

Property Let ButtonSpacingH(ByVal value As Long)
    VCOMObject.ButtonSpacingH = value
End Property

Property Get ButtonSpacingV() As Long
    ButtonSpacingV = VCOMObject.ButtonSpacingV
End Property

Property Let ButtonSpacingV(ByVal value As Long)
    VCOMObject.ButtonSpacingV = value
End Property

Property Get PaddingH() As Long
    PaddingH = VCOMObject.PaddingH
End Property

Property Let PaddingH(ByVal value As Long)
    VCOMObject.PaddingH = value
End Property

Property Get PaddingV() As Long
    PaddingV = VCOMObject.PaddingV
End Property

Property Let PaddingV(ByVal value As Long)
    VCOMObject.PaddingV = value
End Property

Property Get ScreenBorderPaddingV() As Long
    ScreenBorderPaddingV = VCOMObject.ScreenBorderPaddingV
End Property

Property Let ScreenBorderPaddingV(ByVal value As Long)
    VCOMObject.ScreenBorderPaddingV = value
End Property

Property Get ColumnPaddingH() As Long
    ColumnPaddingH = VCOMObject.ColumnPaddingH
End Property

Property Let ColumnPaddingH(ByVal value As Long)
    VCOMObject.ColumnPaddingH = value
End Property

Property Get LineSpacingV() As Long
    LineSpacingV = VCOMObject.LineSpacingV
End Property

Property Let LineSpacingV(ByVal value As Long)
    VCOMObject.LineSpacingV = value
End Property

Property Get MainBackColor() As Long
    MainBackColor = VCOMObject.MainBackColor
End Property

Property Let MainBackColor(ByVal value As Long)
    VCOMObject.MainBackColor = value
End Property

Property Get MainBackColor2() As Long
    MainBackColor2 = VCOMObject.MainBackColor2
End Property

Property Let MainBackColor2(ByVal value As Long)
    VCOMObject.MainBackColor2 = value
End Property

Property Get MainBackFillType() As Long
    MainBackFillType = VCOMObject.MainBackFillType
End Property

Property Let MainBackFillType(ByVal value As Long)
    VCOMObject.MainBackFillType = value
End Property

Property Get MoreInfoBackColor2() As Long
    MoreInfoBackColor2 = VCOMObject.MoreInfoBackColor2
End Property

Property Let MoreInfoBackColor2(ByVal value As Long)
    VCOMObject.MoreInfoBackColor2 = value
End Property

Property Get MoreInfoBackFillType() As Long
    MoreInfoBackFillType = VCOMObject.MoreInfoBackFillType
End Property

Property Let MoreInfoBackFillType(ByVal value As Long)
    VCOMObject.MoreInfoBackFillType = value
End Property

Property Get ButtonBarBackColor() As Long
    ButtonBarBackColor = VCOMObject.ButtonBarBackColor
End Property

Property Let ButtonBarBackColor(ByVal value As Long)
    VCOMObject.ButtonBarBackColor = value
End Property

Property Get ButtonBarBackColor2() As Long
    ButtonBarBackColor2 = VCOMObject.ButtonBarBackColor2
End Property

Property Let ButtonBarBackColor2(ByVal value As Long)
    VCOMObject.ButtonBarBackColor2 = value
End Property

Property Get ButtonBarBackFillType() As Long
    ButtonBarBackFillType = VCOMObject.ButtonBarBackFillType
End Property

Property Let ButtonBarBackFillType(ByVal value As Long)
    VCOMObject.ButtonBarBackFillType = value
End Property

Property Get MaxNumCallStackItems() As Long
    MaxNumCallStackItems = VCOMObject.MaxNumCallStackItems
End Property

Property Let MaxNumCallStackItems(ByVal value As Long)
    VCOMObject.MaxNumCallStackItems = value
End Property

Property Get MaxNumVariablesItems() As Long
    MaxNumVariablesItems = VCOMObject.MaxNumVariablesItems
End Property

Property Let MaxNumVariablesItems(ByVal value As Long)
    VCOMObject.MaxNumVariablesItems = value
End Property

Property Get DefaultButtonIsBold() As Boolean
    DefaultButtonIsBold = VCOMObject.DefaultButtonIsBold
End Property

Property Let DefaultButtonIsBold(ByVal value As Boolean)
    VCOMObject.DefaultButtonIsBold = value
End Property

Property Get ShowMoreInfoButton() As Boolean
    ShowMoreInfoButton = VCOMObject.ShowMoreInfoButton
End Property

Property Let ShowMoreInfoButton(ByVal value As Boolean)
    VCOMObject.ShowMoreInfoButton = value
End Property

Property Get AllowEnterKey() As Boolean
    AllowEnterKey = VCOMObject.AllowEnterKey
End Property

Property Let AllowEnterKey(ByVal value As Boolean)
    VCOMObject.AllowEnterKey = value
End Property

Property Get AllowTabKey() As Boolean
    AllowTabKey = VCOMObject.AllowTabKey
End Property

Property Let AllowTabKey(ByVal value As Boolean)
    VCOMObject.AllowTabKey = value
End Property

Property Get DefaultButtonID() As Long
    DefaultButtonID = VCOMObject.DefaultButtonID
End Property

Property Let DefaultButtonID(ByVal value As Long)
    VCOMObject.DefaultButtonID = value
End Property

Property Get CustomImageTransparentColor() As Long
    CustomImageTransparentColor = VCOMObject.CustomImageTransparentColor
End Property

Property Let CustomImageTransparentColor(ByVal value As Long)
    VCOMObject.CustomImageTransparentColor = value
End Property

Property Get OnKeyDown() As String
    OnKeyDown = VCOMObject.OnKeyDown
End Property

Property Let OnKeyDown(ByVal value As String)
    VCOMObject.OnKeyDown = value
End Property

Property Get FlashWindowOnOpen() As Boolean
    FlashWindowOnOpen = VCOMObject.FlashWindowOnOpen
End Property

Property Let FlashWindowOnOpen(ByVal value As Boolean)
    VCOMObject.FlashWindowOnOpen = value
End Property

Property Get MoreInfoBackColor() As Long
    MoreInfoBackColor = VCOMObject.MoreInfoBackColor
End Property

Property Let MoreInfoBackColor(ByVal value As Long)
    VCOMObject.MoreInfoBackColor = value
End Property

Property Get AllowArrowKeys() As Boolean
    AllowArrowKeys = VCOMObject.AllowArrowKeys
End Property

Property Let AllowArrowKeys(ByVal value As Boolean)
    VCOMObject.AllowArrowKeys = value
End Property

Sub AddButton(ByVal DisplayName As String, ByVal ButtonAction As DefaultButtonActions)
    Call VCOMObject.AddButton(DisplayName, ButtonAction)
End Sub

Sub AddCustomButton(ByVal DisplayName As String, ByVal OnClickProcName As String)
    Call VCOMObject.AddCustomButton(DisplayName, OnClickProcName)
End Sub

Sub RemoveAllButtons()
    Call VCOMObject.RemoveAllButtons
End Sub

Function GetButtonIDByIndex(ByVal Index As Long) As Variant
    Call VCOMObject.AssignVar(GetButtonIDByIndex, VCOMObject.GetButtonIDByIndex(Index))
End Function

Function GetButtonCaptionByIndex(ByVal Index As Long) As Variant
    Call VCOMObject.AssignVar(GetButtonCaptionByIndex, VCOMObject.GetButtonCaptionByIndex(Index))
End Function

Property Get ButtonCount() As Long
    ButtonCount = VCOMObject.ButtonCount
End Property

Property Get Timeout() As Long
    Timeout = VCOMObject.Timeout
End Property

Property Let Timeout(ByVal value As Long)
    VCOMObject.Timeout = value
End Property

Property Get Timedout() As Boolean
    Timedout = VCOMObject.Timedout
End Property

Property Get CustomImageData() As String
    CustomImageData = VCOMObject.CustomImageData
End Property

Property Let CustomImageData(ByVal BitmapFileData As String)
    VCOMObject.CustomImageData = BitmapFileData
End Property

Sub LoadCustomImageFromFile(ByVal filePath As String)
    Call VCOMObject.LoadCustomImageFromFile(filePath)
End Sub

Function InternalCompressData(ByVal Data As String, ByVal CompressorModule As String) As String
    InternalCompressData = VCOMObject.InternalCompressData(Data, CompressorModule)
End Function
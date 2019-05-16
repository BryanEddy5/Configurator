Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =22
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8280
    DatasheetFontHeight =10
    ItemSuffix =39
    Left =10110
    Top =3105
    Right =18390
    Bottom =8610
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x28180a37cb0be540
    End
    Caption ="Create Item"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyPress ="[Event Procedure]"
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
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =5520
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =2400
                    Top =2640
                    Name ="btnSubmit"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =1740
                    Top =90
                    TabIndex =1
                    Name ="BasicCable"
                    AfterUpdate ="mcrLocationBasicCable"

                    LayoutCachedLeft =1740
                    LayoutCachedTop =90
                    LayoutCachedWidth =2000
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =1965
                            Top =60
                            Width =1575
                            Height =240
                            Name ="Label1"
                            Caption ="Double Pass Cable"
                            LayoutCachedLeft =1965
                            LayoutCachedTop =60
                            LayoutCachedWidth =3540
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin OptionButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3780
                    Top =90
                    TabIndex =2
                    Name ="BasicProduct"
                    AfterUpdate ="mcrLocationBasicProduct"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =90
                    LayoutCachedWidth =4040
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =4005
                            Top =60
                            Width =1995
                            Height =240
                            Name ="Label3"
                            Caption ="Single Pass Cable"
                            LayoutCachedLeft =4005
                            LayoutCachedTop =60
                            LayoutCachedWidth =6000
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =4140
                    Top =2640
                    TabIndex =3
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2820
                    Top =750
                    TabIndex =4
                    Name ="Bypass"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =750
                    LayoutCachedWidth =3080
                    LayoutCachedHeight =990
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =3045
                            Top =720
                            Width =2040
                            Height =240
                            Name ="txtBypass"
                            Caption ="Bypass Existing Item Check"
                            LayoutCachedLeft =3045
                            LayoutCachedTop =720
                            LayoutCachedWidth =5085
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3240
                    Top =390
                    TabIndex =5
                    Name ="Armor"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =390
                    LayoutCachedWidth =3500
                    LayoutCachedHeight =630
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =3465
                            Top =360
                            Width =1095
                            Height =240
                            Name ="Label32"
                            Caption ="Armored Item"
                            LayoutCachedLeft =3465
                            LayoutCachedTop =360
                            LayoutCachedWidth =4560
                            LayoutCachedHeight =600
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    Left =1020
                    Top =2580
                    Width =720
                    Height =360
                    TabIndex =6
                    Name ="PI"
                    ControlSource ="=3.1415927"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1020
                    LayoutCachedTop =2580
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =2940
                    ForeThemeColorIndex =0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =360
                    Top =2580
                    Width =600
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label244"
                    Caption ="PI"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =360
                    LayoutCachedTop =2580
                    LayoutCachedWidth =960
                    LayoutCachedHeight =2940
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =2340
                    Top =3480
                    Width =1380
                    TabIndex =7
                    Name ="NewItemNumber"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2340
                    LayoutCachedTop =3480
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =3720
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =420
                    Top =3480
                    Width =1920
                    Height =240
                    Name ="Label224"
                    Caption ="NewItemNumber"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =420
                    LayoutCachedTop =3480
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =3720
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =2340
                    Top =3120
                    Width =1380
                    TabIndex =8
                    Name ="OracleOrg"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2340
                    LayoutCachedTop =3120
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =3360
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =420
                    Top =3120
                    Width =1920
                    Height =240
                    Name ="Label236"
                    Caption ="OracleOrg"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =420
                    LayoutCachedTop =3120
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =3360
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =6000
                    Top =3120
                    Width =1380
                    TabIndex =9
                    Name ="CopyCount"
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6000
                    LayoutCachedTop =3120
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =3360
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =4080
                    Top =3120
                    Width =1920
                    Height =240
                    Name ="Label259"
                    Caption ="CopyCount"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =3120
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =3360
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =6000
                    Top =3480
                    Width =1380
                    TabIndex =10
                    Name ="StdLotSizeNew"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6000
                    LayoutCachedTop =3480
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =3720
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =4080
                    Top =3480
                    Width =1920
                    Height =240
                    Name ="Label261"
                    Caption ="StdLotSizeNew"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =3480
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =3720
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =6000
                    Top =3840
                    Width =1380
                    TabIndex =11
                    Name ="PrimaryUOMCodeNew"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6000
                    LayoutCachedTop =3840
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =4080
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =4080
                    Top =3840
                    Width =1920
                    Height =240
                    Name ="Label263"
                    Caption ="PrimaryUOMCodeNew"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =3840
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =4080
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =2340
                    Top =3840
                    Width =1380
                    TabIndex =12
                    Name ="ItemSetupOrg"
                    DefaultValue ="=\"CAB\""
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2340
                    LayoutCachedTop =3840
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =4080
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =420
                    Top =3840
                    Width =1920
                    Height =240
                    Name ="Label285"
                    Caption ="ItemSetupOrg"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =420
                    LayoutCachedTop =3840
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =4080
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =2340
                    Top =4200
                    Width =1380
                    TabIndex =13
                    Name ="CopyItem"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2340
                    LayoutCachedTop =4200
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =4440
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =420
                    Top =4200
                    Width =1920
                    Height =240
                    Name ="Label289"
                    Caption ="CopyItem"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =420
                    LayoutCachedTop =4200
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =4440
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =2340
                    Top =4560
                    Width =1380
                    TabIndex =14
                    Name ="NewItemType"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2340
                    LayoutCachedTop =4560
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =4800
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =420
                    Top =4560
                    Width =1920
                    Height =240
                    Name ="Label301"
                    Caption ="NewItemType"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =420
                    LayoutCachedTop =4560
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =4800
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =2340
                    Top =4920
                    Width =1380
                    TabIndex =15
                    Name ="FiberType1"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =2340
                    LayoutCachedTop =4920
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =5160
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =420
                    Top =4920
                    Width =1920
                    Height =240
                    Name ="Label311"
                    Caption ="FiberType1"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =420
                    LayoutCachedTop =4920
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =5160
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =6000
                    Top =4200
                    Width =1380
                    TabIndex =16
                    Name ="NewBaseItem"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6000
                    LayoutCachedTop =4200
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =4440
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =4080
                    Top =4200
                    Width =1920
                    Height =240
                    Name ="Label313"
                    Caption ="NewBaseItem"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =4200
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =4440
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =6000
                    Top =4560
                    Width =1380
                    TabIndex =17
                    Name ="NewPrintExtension"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6000
                    LayoutCachedTop =4560
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =4800
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =4080
                    Top =4560
                    Width =1920
                    Height =240
                    Name ="Label315"
                    Caption ="NewPrintExtension"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =4560
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =4800
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =6000
                    Top =4920
                    Width =1380
                    TabIndex =18
                    Name ="ComponentItemNumber"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6000
                    LayoutCachedTop =4920
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =5160
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =4080
                    Top =4920
                    Width =1920
                    Height =240
                    Name ="Label317"
                    Caption ="ComponentItemNumber"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =4920
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =5160
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =6000
                    Top =5280
                    Width =1380
                    TabIndex =19
                    Name ="ItemNumberSelect"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =6000
                    LayoutCachedTop =5280
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =5520
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =4140
                    Top =5280
                    Width =1920
                    Height =240
                    Name ="Label319"
                    Caption ="ItemNumberSelect"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =4140
                    LayoutCachedTop =5280
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =5520
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3300
                    Top =1200
                    Width =3600
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    Name ="ItemNo"
                    Format ="@;\" <enter design code>\""
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    OnKeyDown ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"

                    LayoutCachedLeft =3300
                    LayoutCachedTop =1200
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1440
                            Top =1200
                            Width =1620
                            Height =360
                            FontSize =10
                            Name ="Item No_Label"
                            Caption ="Enter New Item #"
                            EventProcPrefix ="Item_No_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3300
                    Top =1680
                    Width =3600
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =21
                    Name ="txItemRequestNumber"
                    Format ="@;\" <enter request number>\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3300
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =1980
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1440
                            Top =1680
                            Width =1620
                            Height =360
                            FontSize =10
                            Name ="Label38"
                            Caption ="Request #"
                            LayoutCachedLeft =1440
                            LayoutCachedTop =1680
                            LayoutCachedWidth =3060
                            LayoutCachedHeight =2040
                        End
                    End
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
Public Event ItemRequestNumber(value As String)
Dim ArmorCheck As Boolean

Public Sub BoxPrint()
    Dim PrintLine1 As String
    Dim PrintType As String
    Dim strNewItemNo  As String



    If strNewItemNo Like "rbn*" Then
        Exit Sub
    End If
    



    PrintType = Mid(ItemNo, 9, 1)


    If PrintType = "1" Then


        strNewItemNo = NextItemNo(Me.ItemNo, False)
        Forms!frmCreateItem![NewOracle#] = strNewItemNo


        If Me.ItemNo Like "*1MFBOX*" Then
            Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-02BX1"
            Forms!frmCreateItem![PrintLine4].value = "MM/YY 00000 FEET     (every 2 feet, reset counter, 1000ft)"
            Forms!frmCreateItem!Customer = "AFL STD -1000FT"

        
        ElseIf Me.ItemNo Like "*2MFBOX*" Then

            Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-02BX2"
            Forms!frmCreateItem![PrintLine4].value = "MM/YY 00000 FEET     (every 2 feet, reset counter, 2000ft)"
            Forms!frmCreateItem!Customer = "AFL STD -2000FT"

        
        ElseIf Me.ItemNo Like "*3MFBOX*" Then

            Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-02BX3"
            Forms!frmCreateItem![PrintLine4].value = "MM/YY 00000 FEET     (every 2 feet, reset counter, 3000ft)"
            Forms!frmCreateItem!Customer = "AFL STD -3000FT"

        Else

            Call CustomBoxLength(1)
        
        End If
    
        Forms!frmCreateItem![PrintType].value = "1"
        Forms!frmCreateItem!PrintLine1 = "AFL OPTICAL CABLE   1-800-AFL-FIBER"
        Forms!frmCreateItem!Customer = "AFL STANDARD"
    
        Call DissablePrint(True)
    ElseIf PrintType = 2 Then

        Call CustomBoxLength(2)


    End If

End Sub

Private Sub CustomBoxLength(PrintType As Integer)
    Dim intBoxValue As String
    Dim dblBoxExtension As String
    Dim Aresponse As Integer

    Aresponse = MsgBox("Is the print in Meters?", vbYesNo, "Continue")

    intBoxValue = InputBox("How many units of cable are in this boxed item?", "Feet of Cable", 0)
    
    
    If intBoxValue Mod 1000 = 0 Then
        dblBoxExtension = Round(intBoxValue / 1000, 0)
    Else
        dblBoxExtension = Round(intBoxValue / 1000, 2)
    End If
    
    
    If PrintType = 1 Then
        If Aresponse = vbYes Then
            Forms!frmCreateItem![NewOracle#] = Forms!frmCreateItem![NewOracle#] & "-01BX" & dblBoxExtension
        Else
            Forms!frmCreateItem![NewOracle#] = Forms!frmCreateItem![NewOracle#] & "-02BX" & dblBoxExtension
        End If
    
        Forms!frmCreateItem!PrintLine4.Locked = True
    Else
        Forms!frmCreateItem![NewOracle#] = Forms!frmCreateItem![NewOracle#] & "BX" & dblBoxExtension
    End If
    
    If Aresponse = vbYes Then
        Forms!frmCreateItem![PrintLine4].value = "MM/YY 000000 METERS   (reset counter, " & intBoxValue & "m)"
        
    Else
        
        Forms!frmCreateItem![PrintLine4] = "MM/YY 00000 FEET     (every 2 feet, reset counter, " & intBoxValue & "ft)"
        
        
    End If




End Sub

Private Sub DissablePrint(Dissable As Boolean)

    Forms!frmCreateItem!Customer.Enabled = Not Dissable
    Forms!frmCreateItem!PID.Enabled = Not Dissable
    Forms!frmCreateItem!Spec.Enabled = Not Dissable
    Forms!frmCreateItem!CustomerRev.Enabled = Not Dissable
    'Forms!frmCreateItem!PrintLine4.Enabled = False

End Sub

Private Sub lockPrintFields(LockField As Boolean)

    Forms!frmCreateItem!PrintLine1.Locked = LockField
    Forms!frmCreateItem!PrintLine2.Locked = LockField
    Forms!frmCreateItem!PrintLine3.Locked = LockField


End Sub

Private Sub PrintType()
    Dim CableType As String
    Dim CableColor  As String
    Dim NewExt As String
    Dim Aresponse As Integer
    CableType = Left(Me.ItemNo, 1)
    CableColor = Mid(Me.ItemNo, 10, 1)
    Dim strNewItemNo As String



    Select Case CableColor
    Case "C"
        NewExt = "-12"
    Case "B"
        NewExt = "-11"
    Case "A"
        NewExt = "-10"
    Case Else
        NewExt = "-0" & CableColor
    End Select

    'Finds the next Oracle Number and appends it to frmCreateItem
    'DoCmd.OpenForm "frmOracleNextNo"
    strNewItemNo = NextItemNo(Me.ItemNo, False)
    Forms!frmCreateItem![NewOracle#] = strNewItemNo
    
    
    

    'Evaluates if this is a special print or AFL Standard
    'If statement identifying special print
    If Mid(Forms!frmMainMenu!ItemNo, 9, 1) = 2 Or Mid(Forms!frmMainMenu!ItemNo, 9, 1) = "S" Then
    
        Call lockPrintFields(False)
    
        Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-10"
    
        If ArmorCheck And Me.ItemNo Like "l*" Then
            Forms!frmCreateItem![NewOracle#].value = Forms!frmCreateItem![NewOracle#].value & "IA"
        End If
    
        Forms!frmCreateItem!PrintLine4.Locked = False


        If Me.ItemNo Like "*BOX*" Then
            Call BoxPrint
            Exit Sub
        End If
    
    
        'Asks user what UOM the customer wants for print
        Aresponse = MsgBox("Is the print in Meters?", vbYesNo, "Continue")
    
    
    
        If Aresponse = vbYes Then
        
            Forms!frmCreateItem![PrintLine4].value = "MM/YY 000000 METERS"
        
        Else
        
            Forms!frmCreateItem![PrintLine4].value = "MM/YY 000000 FEET     (Print every other foot)"
        
        End If
    

  
    
    ElseIf Mid(Forms!frmMainMenu!ItemNo, 9, 1) = 1 Or Forms!frmMainMenu!ItemNo Like "m*@*" Then
    
        Call lockPrintFields(True)
    
        If Me.ItemNo Like "*BOX*" Then
            Call BoxPrint
        Else
        
            If Me.ItemNo Like "*msha*" Then
                Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-04"
                Forms!frmCreateItem![PrintLine4].value = "MM/YY 000000 FEET     (Print every other foot)"
            Else
                Aresponse = MsgBox("Is the print in Meters?", vbYesNo, "Continue")
                If Aresponse = vbYes Then
                    Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-01"
                    Forms!frmCreateItem![PrintLine4].value = "MM/YY 000000 METERS"
                
                
                Else
                
                    Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-02"
                    Forms!frmCreateItem![PrintLine4].value = "MM/YY 000000 FEET     (Print every other foot)"
                
                
                End If
            End If
    
    
            Forms!frmCreateItem![PrintType].value = "1"
            Forms!frmCreateItem![Customer].value = "AFL STANDARD"
        
            If bFurcation Then
                Forms!frmCreateItem!PrintLine1 = "AFL FURCATION TUBE  1-800-AFL-FIBER"
        
            ElseIf bCenterMember Then
                Forms!frmCreateItem!PrintLine1 = "[Insert Frp OD]mm to [Insert Final OD]mm  [Material Type]"
            Else
                If Left(Me.ItemNo, 1) = "A" Then
                    Forms!frmCreateItem!PrintLine1 = "AFL eABF OPTICAL CABLE"
                Else
            
                    Forms!frmCreateItem!PrintLine1 = "AFL OPTICAL CABLE   1-800-AFL-FIBER"
                End If
            End If

        
            Call DissablePrint(True)
    
        End If
    
    ElseIf Mid(Forms!frmMainMenu!ItemNo, 9, 1) = "U" Then
    
    
    
        If strNewItemNo Like "rbn*" Then
    
        Else
            Select Case CableType
            Case "F", "T"
                Forms!frmCreateItem![NewOracle#].value = strNewItemNo & NewExt
            Case Else
                Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-00"
            End Select
        End If
        Forms!frmCreateItem![PrintType].value = "0"
        Forms!frmCreateItem![Customer].value = "UNPRINTED"
    
        Forms!frmCreateItem!Customer.Enabled = False
        Forms!frmCreateItem!PID.Enabled = False
        Forms!frmCreateItem!Spec.Enabled = False
        Forms!frmCreateItem!CustomerRev.Enabled = False
        Forms!frmCreateItem!ListingCompany = "NONE"
        Forms!frmCreateItem!ListingCompany.Enabled = False
        Forms!frmCreateItem!ListingType.Enabled = False
        Forms!frmCreateItem!ListingTypeCanada.Enabled = False
    

    
    ElseIf Mid(Forms!frmMainMenu!ItemNo, 9, 1) = "T" Then
    
        Forms!frmCreateItem![NewOracle#].value = strNewItemNo & "-TEST"
        Forms!frmCreateItem![PrintType].value = "1"
        Forms!frmCreateItem![Customer].value = "AFL TEST CABLE"
        Forms!frmCreateItem!PrintLine1 = "AFL TEST CABLE"
        Forms!frmCreateItem!PrintLine4 = "MM/YY 000000 METERS"
    
        Forms!frmCreateItem!Customer.Enabled = False
    
        Call lockPrintFields(False)

    
    Else
        MsgBox ("Check design code - missing/incorrect print indicator")
        DoCmd.Close acForm, "frmCreatItem"
        'Me.Visible = True
        DoCmd.Close acForm, "frmOracleNextNo"
        If CurrentProject.AllForms("frmCreateItem").IsLoaded Then
            Call Forms("frmCreateItem").btnClose_Click
        End If
        If CurrentProject.AllForms("frmFiberMatchExisting").IsLoaded Then
            Call Forms("frmFiberMatchExisting").btnClose_Click
        End If
    
    End If
    DoCmd.Close acForm, "frmOracleNextNo"



End Sub

Private Sub MixedFiber()

    If Me.ItemNo Like "*/*" Then
        Forms!frmCreateItem!MixedFiber = 1
        Forms!frmCreateItem!Fiber2.Visible = True
        Forms!frmCreateItem!OM2.Visible = True
        Forms!frmCreateItem!BIF2.Visible = True
        Forms!frmCreateItem!SubX = 1
        Forms!frmCreateItem!FiberSpecs2.Visible = True
    Else
        Forms!frmCreateItem!MixedFiber = 0
        Forms!frmCreateItem!Fiber2.Visible = False
        Forms!frmCreateItem!Fiber2 = "NONE"
        Forms!frmCreateItem!OM2.Visible = False
        Forms!frmCreateItem!BIF2.Visible = False
        Forms!frmCreateItem!FiberSpecs2.Visible = False
    End If

End Sub

Public Sub FiberExt()


    Forms!frmCreateItem!FiberExt = FiberExtension(Me.[ItemNo])

End Sub

Private Sub Combo2_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[Item No] = '" & Me![Combo2] & "'"
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_cmdCloseForm_Click

    'OpenAllDatabases False
 

    DoCmd.Close
 
Exit_cmdCloseForm_Click:
    Exit Sub
 
Err_cmdCloseForm_Click:
    MsgBox Err.Description
    Resume Exit_cmdCloseForm_Click

End Sub

Private Sub btnDesignCode_Click()
    'Me.FormFooter.Visible = True
End Sub

Public Sub btnSubmit_Click()
    Dim lResponse As Integer
    Dim KResponse As Integer
    Dim SubColor As String
    Dim cTBType As String



    On Error GoTo Error_Handler:

    cTBType = Mid(Forms!frmMainMenu!ItemNo, 12, 1)
    SubColor = Mid(Me.ItemNo, 11, 1)




    DoCmd.Hourglass True



    If IsNull(Me.ItemNo) = True Then
        MsgBox ("Please enter a design code")
        DoCmd.Hourglass False
        Exit Sub
    End If
    
    If CurrentProject.AllForms("frmCreateItem").IsLoaded Then
        DoCmd.Close acForm, "frmCreateItem"
    End If
    
    
    
    DoCmd.OpenForm "frmBaseExist"
    
    If IsNull(Forms!frmBaseExist!DesignCode) = True And Not Me.ItemNo Like "l*" And Not Me.ItemNo Like "*m????@*" And Mid(Me.ItemNo, 9, 1) <> "T" Then
        'If FindExistingBase = False Then
        KResponse = MsgBox("No active Base or incorrect Tight Buffer type." & vbCrLf & "Please contact Product Engineering", vbCritical)
        DoCmd.Close acForm, "frmBaseExist"
        DoCmd.Hourglass False
        Exit Sub
    End If
        

    DoCmd.Close acForm, "frmBaseExist"
        
        
    'Idenfities an amored build and changes to selection to Basic Cable
    If Me.ItemNo Like "*-AIA*" Then
        ArmorCheck = True
    End If
        
        

        
    lResponse = MsgBox("Are you ready to create this Item", vbYesNo, "Continue")
        
    If lResponse = vbNo Then
        DoCmd.Hourglass False
        Exit Sub
    End If
        
    'If lResponse = vbYes Then
    Forms!frmMainMenu.Form.Visible = False
            
    DoCmd.OpenForm "frmCreateItem"
            
            
    'Set rst = dbs.OpenRecordset("qryMainMenuMatch", dbOpenDynaset)
    If ArmorCheck Then
        DoCmd.OpenForm "frmOracleMatch"
                
        'If Forms!frmOracleMatch.Exists = -1 Then 'commented 6 lines of code for new armoring logic (unmatching jacket colors)
        Forms!frmCreateItem!Armor = -1
        Forms!frmCreateItem!Exists = -1
        Forms!frmMainMenu!Bypass = -1
        DoCmd.OpenForm "frmFiberMatchExisting"
        'Else
        'Forms!frmCreateItem!Exists = 0
        'Forms!frmMainMenu!Bypass = 0
        'End If
    Else
        Forms!frmCreateItem!Exists = 0
        Forms!frmMainMenu!Bypass = 0
        Forms!frmCreateItem!Armor = 0
    End If
            

    If CurrentProject.AllForms("frmCreateItem").IsLoaded Then
            
        If ArmorCheck Then
            Call ArmorSelection(Me.ItemNo)       'Inputs armor print and type for item to be built
        End If
            
        Call PrintType                           'Asks user what UOM the customer wants for print
        Call FiberExt                            'Identifies if the Item has a fiber Extension indicating a certain fiber to use
        Call MixedFiber                          'Identifies mixed fiber cables
            
        'If Forms!frmCreateItem!Armor <> -1 Then
        Call Forms("frmCreateItem").Autofill     'Chooses default information for item build
        RaiseEvent ItemRequestNumber(Nz(Me.txItemRequestNumber, ""))
        'End If
    End If
            
    'End If
    'End If


    DoCmd.Close acForm, "frmOracleMatch"

    DoCmd.Hourglass False

Exit_Error_Handler:
    Exit Sub

Error_Handler:
    DoCmd.Hourglass False
    MsgBox Err.Description
    If CurrentProject.AllForms("frmCreateItem").IsLoaded Then
        Call Forms("frmCreateItem").btnClose_Click
    End If
    If CurrentProject.AllForms("frmFiberMatchExisting").IsLoaded Then
        'Call Forms("frmFiberMatchExisting").btnClose_Click
        DoCmd.Close acForm, "frmFiberMatchExisting"
    End If
    DoCmd.Close acForm, "frmBaseExist"
    DoCmd.Close acForm, "frmOracleNextNo"
    'Me.Visible = True
    '    DoCmd.Close acForm, ""
    Resume Exit_Error_Handler

    


End Sub

Private Sub Command28_Click()
    On Error GoTo Err_cmdCloseForm_Click
 
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close
 
Exit_cmdCloseForm_Click:
    Exit Sub
 
Err_cmdCloseForm_Click:
    MsgBox Err.Description
    Resume Exit_cmdCloseForm_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Me.Refresh
        Call btnSubmit_Click
        Beep
    End If
End Sub

Private Sub Form_Load()
    Bypass.value = 0
    BasicProduct.value = -1
    BasicCable.value = 0
    Me.ItemNo.SetFocus
    
    'OpenAllDatabases True
    
    
End Sub

Private Sub frmCreateItem_Hold()
    Dim TBType As String

    TBType = Mid(Me.ItemNo, 12, 1)

    Forms!frmCreateItem!TBType = TBType

End Sub

Private Sub ItemNo_AfterUpdate()
    If Not IsNull(Me.ItemNo) Then
        Me.ItemNo = Replace(Me.ItemNo, Chr(13) & Chr(10), "")
    End If
End Sub

Private Sub ItemNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.Refresh
        Call btnSubmit_Click
        Beep
    End If
End Sub

Private Sub ItemNo_KeyPress(KeyAscii As Integer)
    Call LimitKeyPress(Me.ItemNo, 30, KeyAscii)
    Call LimitChange(Me.ItemNo, 30)
End Sub

Sub LimitKeyPress(ctl As control, iMaxLen As Integer, KeyAscii As Integer)
    On Error GoTo Err_LimitKeyPress
    ' Purpose:  Limit the text in an unbound text box/combo.
    ' Usage:    In the control's KeyPress event procedure:
    '             Call LimitKeyPress(Me.MyTextBox, 12, KeyAscii)
    ' Note:     Requires LimitChange() in control's Change event also.

    If Len(ctl.Text) - ctl.SelLength >= iMaxLen Then
        If KeyAscii <> vbKeyBack Then
            KeyAscii = 0
            Beep
        End If
    End If

Exit_LimitKeyPress:
    Exit Sub

Err_LimitKeyPress:
    'Call LogError(Err.Number, Err.Description, "LimitKeyPress()")
    Resume Exit_LimitKeyPress
End Sub

Sub LimitChange(ctl As control, iMaxLen As Integer)
    On Error GoTo Err_LimitChange
    ' Purpose:  Limit the text in an unbound text box/combo.
    ' Usage:    In the control's Change event procedure:
    '               Call LimitChange(Me.MyTextBox, 12)
    ' Note:     Requires LimitKeyPress() in control's KeyPress event also.

    If Len(ctl.Text) > iMaxLen Then
        MsgBox "Truncated to " & iMaxLen & " characters.", vbExclamation, "Too long"
        ctl.Text = Left(ctl.Text, iMaxLen)
        ctl.SelStart = iMaxLen
    End If

Exit_LimitChange:
    Exit Sub

Err_LimitChange:
    'Call LogError(Err.Number, Err.Description, "LimitChange()")
    Resume Exit_LimitChange
End Sub

Sub LoadTables()

    'Clear Tables
    DoCmd.OpenQuery "qryCLEARBasicCableSheath"
    DoCmd.OpenQuery "qryCLEARBasicProductConstruction"
    DoCmd.OpenQuery "qryCLEARUnitSeries"

    'Load Tables
    DoCmd.OpenQuery "qryAPPENDBasicCableSheath"
    DoCmd.OpenQuery "qryAPPENDBasicProductConstruction"
    DoCmd.OpenQuery "qryAPPENDUnitSeries"


End Sub

Function FindExistingBase() As Boolean

    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim strSQL As String

    Set cnn = CurrentProject.connection

    strSQL = "SELECT * FROM qryItemNoSearchMatchBase"

    rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic


    If rst.EOF = True Then
        MsgBox ("No Existing Base")
        FindExistingBase = False
    Else
        FindExistingBase = False
    End If
    
    rst.Close
    cnn.Close


End Function

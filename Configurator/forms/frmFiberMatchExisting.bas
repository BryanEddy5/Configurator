Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    AllowUpdating =1
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =26940
    DatasheetFontHeight =10
    ItemSuffix =51
    Left =5235
    Top =885
    Right =28215
    Bottom =7290
    DatasheetGridlinesColor =12632256
    Filter ="[Fiber2] = 'NONE' AND [Fiber] = 'FBR00374' AND [Base] ='LQ144-3.0:CCC' AND [Colo"
        "r] = 'YELLOW' AND RibbonHighCure = False  AND left([Item No],8) = 'LQ144930'  AN"
        "D mid([Item No],10,3) = '99B' AND mid([Item No],9,1) = 'U'"
    RecSrcDt = Begin
        0x7eacf40a740fe540
    End
    RecordSource ="qryMatchFibertoExisting"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
        Begin FormHeader
            Height =2040
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =93
                    Left =11700
                    Top =1680
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text7"
                    Caption ="Fiber"
                    LayoutCachedLeft =11700
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =1140
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Oracle"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =15180
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text9"
                    Caption ="Jacket"
                    LayoutCachedLeft =15180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =20760
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text8"
                    Caption ="Color"
                    LayoutCachedLeft =20760
                    LayoutCachedTop =1680
                    LayoutCachedWidth =22620
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =17100
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text10"
                    Caption ="Base"
                    LayoutCachedLeft =17100
                    LayoutCachedTop =1680
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =215
                    Left =11700
                    Top =1680
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label15"
                    Caption ="Fiber 1"
                    LayoutCachedLeft =11700
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =7320
                    Top =1680
                    Width =2520
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label17"
                    Caption ="Customer"
                    LayoutCachedLeft =7320
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =22740
                    Top =1680
                    Width =2220
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label19"
                    Caption ="Print Line 4"
                    LayoutCachedLeft =22740
                    LayoutCachedTop =1680
                    LayoutCachedWidth =24960
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =2880
                    Top =1680
                    Width =3000
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Catalog Number"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =6000
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label24"
                    Caption ="Active"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =1980
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontFamily =18
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Top =540
                    Width =6060
                    Height =480
                    ColumnOrder =0
                    FontSize =16
                    FontWeight =500
                    Name ="ItemNo"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            SpecialEffect =2
                            OverlapFlags =93
                            TextFontFamily =18
                            Left =240
                            Top =540
                            Width =1620
                            Height =300
                            FontSize =16
                            FontWeight =500
                            Name ="Label32"
                            Caption ="New Item:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =13440
                    Top =1680
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label33"
                    Caption ="Fiber 2"
                    LayoutCachedLeft =13440
                    LayoutCachedTop =1680
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =180
                    Top =420
                    Width =7920
                    Name ="Box35"
                End
                Begin Line
                    OverlapFlags =119
                    Left =1800
                    Top =420
                    Width =0
                    Height =720
                    Name ="Line36"
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =18780
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label38"
                    Caption ="Color Chip"
                    LayoutCachedLeft =18780
                    LayoutCachedTop =1680
                    LayoutCachedWidth =20640
                    LayoutCachedHeight =1980
                End
                Begin Label
                    SpecialEffect =2
                    OverlapFlags =85
                    Left =9900
                    Top =1680
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label46"
                    Caption ="PID"
                    LayoutCachedLeft =9900
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =1980
                End
            End
        End
        Begin Section
            Height =486
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =11700
                    Top =120
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="Text0"
                    ControlSource ="Fiber"

                    LayoutCachedLeft =11700
                    LayoutCachedTop =120
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =20760
                    Top =120
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="Color"

                    LayoutCachedLeft =20760
                    LayoutCachedTop =120
                    LayoutCachedWidth =22620
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15180
                    Top =120
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="Text3"
                    ControlSource ="Jacket"

                    LayoutCachedLeft =15180
                    LayoutCachedTop =120
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17100
                    Top =120
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="Text4"
                    ControlSource ="Base"

                    LayoutCachedLeft =17100
                    LayoutCachedTop =120
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =120
                    Width =1620
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="Text6"
                    ControlSource ="Oracle"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =120
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =180
                    Top =60
                    Width =846
                    Height =366
                    TabIndex =5
                    Name ="btnSelect"
                    Caption ="New Ext"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1026
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =11700
                    Top =120
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    Name ="Text16"
                    ControlSource ="Fiber"

                    LayoutCachedLeft =11700
                    LayoutCachedTop =120
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7320
                    Top =120
                    Width =2520
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =7
                    Name ="Text18"
                    ControlSource ="Customer"

                    LayoutCachedLeft =7320
                    LayoutCachedTop =120
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =22740
                    Top =120
                    Width =3540
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =8
                    Name ="Text20"
                    ControlSource ="PrintLine4"

                    LayoutCachedLeft =22740
                    LayoutCachedTop =120
                    LayoutCachedWidth =26280
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =120
                    Width =3000
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =9
                    Name ="Text22"
                    ControlSource ="Item No"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =120
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6000
                    Top =120
                    Width =660
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =11
                    Name ="Text25"
                    ControlSource ="active"
                    Format ="Yes/No"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =120
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6780
                    Top =60
                    Width =426
                    Height =426
                    TabIndex =12
                    Name ="btnFind"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadad000dadada ,
                        0xadadad0bb00dadaddadada0bb00adadaadadada00dadadaddadada0b00dadada ,
                        0xadadad0b00adadaddadada0b00dadadaadadad0bb00dadaddada00d0bb00dada ,
                        0xada0b00d0bb00daddad0b00ad0b00adaada0bb000bb00daddada0bbbbb00dada ,
                        0xadada000000dadad000000000000000000000000000000000000000000000000 ,
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

                    LayoutCachedLeft =6780
                    LayoutCachedTop =60
                    LayoutCachedWidth =7206
                    LayoutCachedHeight =486
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =26340
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =14
                    Name ="Text30"
                    ControlSource ="Location"

                    LayoutCachedLeft =26340
                    LayoutCachedTop =120
                    LayoutCachedWidth =26700
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13440
                    Top =120
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =15
                    Name ="Text34"
                    ControlSource ="Fiber2"

                    LayoutCachedLeft =13440
                    LayoutCachedTop =120
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18780
                    Top =120
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =16
                    Name ="Text39"
                    ControlSource ="ColorChip"

                    LayoutCachedLeft =18780
                    LayoutCachedTop =120
                    LayoutCachedWidth =20640
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =26280
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =17
                    Name ="Text40"
                    ControlSource ="PrintLine3"

                    LayoutCachedLeft =26280
                    LayoutCachedTop =120
                    LayoutCachedWidth =26640
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =26400
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =18
                    Name ="Text41"
                    ControlSource ="PrintLine2"

                    LayoutCachedLeft =26400
                    LayoutCachedTop =120
                    LayoutCachedWidth =26760
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =26580
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =19
                    Name ="Text42"
                    ControlSource ="PrintLine1"

                    LayoutCachedLeft =26580
                    LayoutCachedTop =120
                    LayoutCachedWidth =26940
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =26580
                    Top =120
                    Width =285
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =20
                    Name ="Text44"
                    ControlSource ="RibbonHighCure"

                    LayoutCachedLeft =26580
                    LayoutCachedTop =120
                    LayoutCachedWidth =26865
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =180
                    Top =60
                    Width =786
                    Height =366
                    TabIndex =10
                    Name ="Armor"
                    Caption ="Armor"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =966
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9900
                    Top =120
                    Width =1680
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =21
                    Name ="Text48"
                    ControlSource ="PID"

                    LayoutCachedLeft =9900
                    LayoutCachedTop =120
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =25920
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =13
                    Name ="Location"
                    ControlSource ="Location"

                    LayoutCachedLeft =25920
                    LayoutCachedTop =120
                    LayoutCachedWidth =26280
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =26040
                    Top =120
                    Width =360
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =22
                    Name ="Text49"
                    ControlSource ="Unit Series"

                    LayoutCachedLeft =26040
                    LayoutCachedTop =120
                    LayoutCachedWidth =26400
                    LayoutCachedHeight =420
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
                    OverlapFlags =93
                    Left =180
                    Top =660
                    Width =1740
                    Name ="btnFiberClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =1920
                    Top =660
                    Width =1800
                    TabIndex =1
                    Name ="btnClose"
                    Caption ="Back to Item"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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
Dim nonDefaultCore As Boolean
Dim SubColor As String, JacketColor As String

Private Sub Armor_Click()

    On Error GoTo Error_Handler:


    'Creating armored item from existing cable

    If Me.Oracle Like "*ia*" And Not Forms!frmCreateItem!ItemNo Like "*aia*" Then
    
    
        MsgBox ("The print indicators in the design code do not match.")
   
        Exit Sub
    
    End If
    
    If Mid(Me.[Item No], 9, 1) = Mid(Forms!frmCreateItem!ItemNo, 9, 1) Then
    
    
    
       
        'Forms!frmCreateItem!ColorChip = Me.ColorChip
    
        DoCmd.OpenForm "frmArmorMatch"
        Forms!frmCreateItem!ArmorNo = Forms!frmArmorMatch!ArmorNo
        Forms!frmCreateItem!Base = Forms!frmArmorMatch!OracleBase
        Forms!frmCreateItem!ArmorCoreOD = Forms!frmArmorMatch!NomOD
        Forms!frmCreateItem!NomOD = Forms!frmArmorMatch!MaxJacketOD
        
        If Me.Base = "GQ012-3.0A" Then
            Forms!frmCreateItem!ArmorNo = 1
            Forms!frmCreateItem!Base = "IA01-AP PVC"
        End If
        
        
        If Me.Customer Like "*afl st*" Then
        
            Forms!frmCreateItem!PrintType = 1
            Forms!frmCreateItem!Customer = Me.Customer
        
        ElseIf Me.Customer Like "*afl gen*" Then
            Forms!frmCreateItem!PrintType = 3
            Forms!frmCreateItem!Customer = Me.Customer
            If IsNull(Forms!frmCreateItem!PID) = False Then
                Forms!frmCreateItem!PID = Me.PID
            End If
            
        ElseIf Me.Customer Like "*unprint*" Then
            Forms!frmCreateItem!PrintType = 0
            Forms!frmCreateItem!Customer = Me.Customer
            
        Else
            Forms!frmCreateItem!PrintType = 2
            Forms!frmCreateItem!Customer = Me.Customer
            
        End If
        
        'Mixed Fiber
        If Forms!frmCreateItem!MixedFiber = 1 Then
            Forms!frmCreateItem!FiberType2 = Me.Fiber2
            Forms!frmCreateItem!Fiber2 = Me.Fiber2
            Forms!frmCreateItem!Fiber2.Enabled = False
            
            DoCmd.OpenForm "frmFiberMode2"
    
            Forms!frmCreateItem!OM2 = Forms!frmFiberMode2![OM#]
            Forms!frmCreateItem!BIF2 = Forms!frmFiberMode2!BIF
        
            DoCmd.Close acForm, "frmFiberMode2"
            
            
        End If
        
        'if core selected is not a default (matching jacket color) then don't update the oracle item number
        If nonDefaultCore = True Then
            Forms!frmCreateItem![NewOracle#] = Forms!frmCreateItem![NewOracle#] & "IA"
        Else
            Forms!frmCreateItem![NewOracle#] = Me.Oracle & "IA"
            Forms!frmCreateItem!Color = Me.Color
        End If
        Forms!frmCreateItem!Armor = -1
        Forms!frmCreateItem!Fiber1 = Me.Fiber
        Forms!frmCreateItem!Fiber = Me.Fiber
        Forms!frmCreateItem!Fiber1.Enabled = False
        Forms!frmCreateItem![Exists].value = 0
        Forms!frmCreateItem!btnUnitSeries.Visible = False
        
        Forms!frmCreateItem!Color.SetFocus
        Forms!frmCreateItem!Base.Enabled = False
        Forms!frmCreateItem!Jacket.Enabled = False
        Forms!frmCreateItem!Color.Enabled = False
        
        DoCmd.OpenForm "frmFiberMode"
    
        Forms!frmCreateItem!OM = Forms!frmFiberMode![OM#]
        Forms!frmCreateItem!BIF = Forms!frmFiberMode!BIF
        
        'DoCmd.OpenForm "frmFiberMatchingExistingPrintLine_1_2"
        
        If Mid(Me.[Item No], 9, 1) = 1 Then
            Forms("frmCreateItem").Fiber1_AfterUpdate
        End If
        'Else
        Forms!frmCreateItem!PrintLine2 = Me.PrintLine2
        Forms!frmCreateItem!PrintLine3 = Me.PrintLine3
        DoCmd.OpenForm "frmFiberMode"
        Forms!frmCreateItem!OM = Forms!frmFiberMode![OM#]
        Forms!frmCreateItem!BIF = Forms!frmFiberMode!BIF
        DoCmd.Close acForm, "frmFiberMode"
        'Else
        Forms("frmCreateItem").ListingAutoFill
        'End If
        
        If IsNull(Forms!frmCreateItem!PrintLine3) = False Then
            Forms!frmCreateItem!PrintLine3 = Replace(Me.PrintLine3, "OFN", "OFC")
        End If
        
        
        Forms!frmCreateItem!PrintLine1 = Me.PrintLine1
        Forms!frmCreateItem!PrintLine4 = Me.PrintLine4
        
        
        DoCmd.Close acForm, "frmArmorMatch"
        'DoCmd.Close acForm, "frmFiberMatchingExistingPrintLine_1_2"
        DoCmd.Close acForm, "frmFiberMatchExisting"
        
        Forms!frmCreateItem!cbSubunit.Enabled = False
        Forms!frmCreateItem!Color.Requery
        Forms!frmCreateItem!Color = Forms!frmCreateItem!Color.ItemData(1)
        Call Forms("frmCreateItem").PrintType_AfterUpdate
        
        Call Forms("frmCreateItem").ListingType_AfterUpdate
        'Call Forms("frmCreateItem").ListingUpdate
        
    End If
 




Error_Exit:
    Exit Sub

Error_Handler:
    If Err.Number = 2113 Then
        MsgBox "Ensure the armor suffix is correct." & vbCrLf & "Example: -AIAE is not a correct armor suffix", vbCritical
    Else
        MsgBox (Err.Description & " " & Err.Number)
    End If

    If CurrentProject.AllForms("frmArmorMatch").IsLoaded Then
        DoCmd.Close acForm, "frmArmorMatch"
    End If
    Resume Error_Exit:

End Sub

Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.name
    'Call Forms("frmCreateItem").btnClose_Click
End Sub

Private Sub btnFiberClose_Click()

    
    Forms!frmCreateItem![Exists].value = 0
    DoCmd.Close acForm, Me.name
    
    If Forms!frmCreateItem!ItemNo Like "*ia*" Then
        MsgBox "No active core was found", vbCritical
        DoCmd.Close acForm, "frmCreateItem"
        Forms!frmMainMenu.Visible = True
    Else
        Call Forms("frmCreateItem").CreateItem
    End If


End Sub

Private Sub btnFind_Click()


    Call OpenAClient(0, Me.Oracle, Nz(Me.Location, 2))


End Sub

Private Sub FilterForSubunitJacketColor()
    Dim Prefix As String
    Dim tightBuffer As String
    tightBuffer = Mid(Forms!frmMainMenu!ItemNo, 12, 1)
    Prefix = Left(Forms!frmMainMenu!ItemNo, 9)
    nonDefaultCore = True
    
    Me.Filter = "left([Item No],12) like '" & Prefix & SubColor & "?" & tightBuffer & "' AND [Oracle] NOT LIKE '*-??IA'"
    Me.FilterOn = True
    Me.Requery
    
    If Not IsNull(Me.Oracle) Then
        Beep
        MsgBox "Core item found where core jacket matches new item's subunit color"
    End If
End Sub

Private Sub btnSelect_Click()
    Dim strOracle As String

    strOracle = Forms!frmCreateItem![NewOracle#]

    DoCmd.SetWarnings False

    If Not Me.Oracle Like "*ia*" And Not Forms!frmCreateItem!ItemNo Like "*aia*" Then
    
        DoCmd.OpenForm "frmExistingNextOracle"
    
        If Forms!frmCreateItem!ItemNo Like "*BOX*" Then
        
            If Mid(Forms!frmCreateItem!ItemNo, 9, 1) <> "2" Then
                Forms!frmCreateItem![NewOracle#] = Left(Me.Oracle, 8) & Mid(strOracle, 9, Len(strOracle) - 8)
            ElseIf Mid(Forms!frmCreateItem!ItemNo, 9, 1) = "2" Then
                If Mid(Forms!frmExistingNextOracle!Oracle, 9, 1) = "0" Then
                    Forms!frmCreateItem![NewOracle#] = Left(Forms!frmExistingNextOracle!Oracle, 8) & "10" & Mid(strOracle, 11, Len(strOracle) - 8)
                Else
                    Forms!frmCreateItem![NewOracle#] = Left(Forms!frmExistingNextOracle!Oracle, 8) & Mid(Forms!frmExistingNextOracle!Oracle, 9, 2) + 1 & Mid(strOracle, 11, Len(strOracle) - 8)
                End If
            End If
        Else
    
            If Mid(Forms!frmCreateItem!ItemNo, 9, 1) = 2 Or Forms!frmCreateItem!PrintType = 2 Then
                If Mid(Forms!frmExistingNextOracle!Oracle, 9, 1) = "0" Then
                    Forms!frmCreateItem![NewOracle#] = Left(Forms!frmExistingNextOracle!Oracle, 8) & "10"
                Else
                    Forms!frmCreateItem![NewOracle#] = Left(Forms!frmExistingNextOracle!Oracle, 8) & Mid(Forms!frmExistingNextOracle!Oracle, 9, 2) + 1
                End If
            Else
                Forms!frmCreateItem![NewOracle#] = Left(Forms!frmExistingNextOracle!Oracle, 8) & Right(Forms!frmCreateItem![NewOracle#], 2)
        
        
            End If
        End If
    
        Forms!frmCreateItem!UnitSeries = Me.[Unit Series]

    

    
        DoCmd.Close acForm, "frmExistingNextOracle"
        Forms!frmCreateItem!Exists = -1
    

    
        DoCmd.Close acForm, "frmFiberMatchExisting"
        Call Forms("frmCreateItem").CreateItem

    Else

    End If

    DoCmd.SetWarnings True




End Sub

Private Sub Form_Load()

    SubColor = Mid(Forms!frmMainMenu!ItemNo, 11, 1)
    JacketColor = Mid(Forms!frmMainMenu!ItemNo, 10, 1)
    'Me.RecordSource = "qryMatchFibertoExisting"

    If sUser = 1 Then
        Me.btnFiberClose.Visible = True
    Else
        Me.btnFiberClose.Visible = False
    End If

    Me.move 0, 0
    If Forms!frmCreateItem!ItemNo Like "*aia*" Then
        Me.btnSelect.Visible = False
        Me.Armor.Visible = True
        Me.btnClose.Visible = False
        Me.btnFiberClose.Caption = "Close"
    Else
        Me.btnSelect.Visible = True
        Me.Armor.Visible = False
        Me.btnFiberClose.Visible = True
        Me.btnClose.Visible = True
        Me.btnFiberClose.Caption = "Create New Parent"
        
        
    End If

    If Forms!frmCreateItem!Armor = -1 Then

        Call Forms("frmMainMenu").FiberExt
       
        
    
    
        If Forms!frmCreateItem!MixedFiber <> 1 Then
        
        
            Select Case Forms!frmCreateItem!FiberExt
            Case 1
                Me.RecordSource = "qryMatchFibertoExistingExtension"
            Case 2
                Me.RecordSource = "qryMatchFibertoExistingBIF"
            Case 3
                Me.RecordSource = "qryMatchFibertoExistingITU"
            Case Else
                Me.RecordSource = "qryMatchFibertoExisting"
            End Select
        Else
        
            Select Case Forms!frmCreateItem!FiberExt
            
            End Select
        End If
        Me.Filter = "left([Item No],12) = '" & Left(Forms!frmMainMenu!ItemNo, 12) & "'" ' AND Oracle NOT LIKE '%-??IA'"
        Me.FilterOn = True
        Me.Requery

    End If



    'If Forms!frmCreateItem![NewOracle#] Like "pt*" Then
    '    Me.RecordSource = "qryMatchFibertoExistingTightBuffers"
    'End If


    Me.ItemNo = Forms!frmCreateItem!ItemNo

    If IsNull(Me.Oracle) = True Then
    
        If Forms!frmCreateItem!Armor = -1 And Forms!frmCreateItem!ItemNo Like "l*" Then
            Exit Sub
        ElseIf Forms!frmCreateItem!Armor <> -1 Then
            Call btnFiberClose_Click
        ElseIf JacketColor <> SubColor Then
            Call FilterForSubunitJacketColor
            If IsNull(Me.Oracle) = True Then
                Call btnFiberClose_Click
            End If
        Else
            Call btnFiberClose_Click
        End If
    End If


End Sub

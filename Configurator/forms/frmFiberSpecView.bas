Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =22
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15000
    DatasheetFontHeight =10
    ItemSuffix =23
    Left =6615
    Top =3465
    Right =21870
    Bottom =6240
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x17f68df5efb9e440
    End
    RecordSource ="qryfrmCreateItemFiberSpec2"
    Caption ="qryfrmCreateItemFiberSpec"
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
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =480
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =4680
                    Top =60
                    Width =1500
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Fiber Oracle item_Label"
                    Caption ="Fiber Oracle item"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Fiber_Oracle_item_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =6480
                    Top =60
                    Width =1200
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="1st Req Freq_Label"
                    Caption ="1st Req Freq"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl1st_Req_Freq_Label"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    Left =7800
                    Top =60
                    Width =1200
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="1st Max Atten_Label"
                    Caption ="1st Max Atten"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl1st_Max_Atten_Label"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    TextFontFamily =18
                    Left =9000
                    Top =60
                    Width =1530
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="1st Min BandW_Label"
                    Caption ="1st Min BandW"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl1st_Min_BandW_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =10710
                    Top =60
                    Width =1050
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="2nd Req Freq_Label"
                    Caption ="2nd Req Freq"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl2nd_Req_Freq_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =12120
                    Top =60
                    Width =1260
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="2nd Max Atten_Label"
                    Caption ="2nd Max Atten"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl2nd_Max_Atten_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =13500
                    Top =60
                    Width =1440
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="2nd Min BandW_Label"
                    Caption ="2nd Min BandW"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl2nd_Min_BandW_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =420
                    Top =60
                    Width =3720
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label16"
                    Caption ="Fiber Description"
                    FontName ="Times New Roman"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =480
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    CanShrink = NotDefault
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6420
                    Top =60
                    Width =1260
                    ColumnWidth =1530
                    ColumnOrder =2
                    FontSize =10
                    FontWeight =700
                    Name ="1stReqFreq"
                    ControlSource ="1st Req Freq"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl1stReqFreq"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7800
                    Top =60
                    Width =1080
                    ColumnWidth =1530
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="1stMaxAtten"
                    ControlSource ="1st Max Atten"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl1stMaxAtten"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9000
                    Top =60
                    Width =1470
                    ColumnWidth =1530
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="1stMinBandW"
                    ControlSource ="1st Min BandW"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl1stMinBandW"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =10710
                    Top =60
                    Width =1350
                    ColumnWidth =1530
                    ColumnOrder =5
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="2ndReqFreq"
                    ControlSource ="2nd Req Freq"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl2ndReqFreq"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =12120
                    Top =60
                    Width =1230
                    Height =255
                    ColumnWidth =1530
                    ColumnOrder =6
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="2ndMaxAtten"
                    ControlSource ="2nd Max Atten"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl2ndMaxAtten"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =13500
                    Top =60
                    ColumnWidth =1530
                    ColumnOrder =7
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="2ndMinBandW"
                    ControlSource ="2nd Min BandW"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Ctl2ndMinBandW"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =420
                    Top =60
                    Width =4200
                    Height =255
                    ColumnWidth =6105
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="FiberType"
                    ControlSource ="Fiber Type"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4740
                    Top =60
                    Width =1560
                    Height =255
                    ColumnOrder =1
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="Fiber"
                    ControlSource ="Fiber Oracle item"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            Height =1008
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1920
                    Top =60
                    Width =960
                    Height =420
                    FontSize =10
                    FontWeight =700
                    Name ="Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Times New Roman"

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

Private Sub btnSaveFiber_Click()
    On Error GoTo Err_btnSaveFiber_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_btnSaveFiber_Click:
    Exit Sub

Err_btnSaveFiber_Click:
    MsgBox Err.Description
    Resume Exit_btnSaveFiber_Click
    
End Sub

Private Sub Close_Click()
    On Error GoTo Err_Close_Click


    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
    MsgBox Err.Description
    Resume Exit_Close_Click
    
End Sub

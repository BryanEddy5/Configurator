Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =16320
    DatasheetFontHeight =10
    ItemSuffix =94
    Left =5955
    Top =2445
    Right =22530
    Bottom =9690
    DatasheetGridlinesColor =12632256
    Filter ="((qryItemInterfaceQTSView.ElementName Like \"FIBER1*\"))"
    OrderBy ="qryItemInterfaceQTSView.ElementName"
    RecSrcDt = Begin
        0x8d7243665352e440
    End
    RecordSource ="qryItemInterfaceQTSView"
    Caption ="QTS Inputs"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    AllowLayoutView =0
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
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
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
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
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
            CanGrow = NotDefault
            Height =1980
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Top =1680
                    Width =480
                    Height =300
                    FontSize =10
                    Name ="Label15"
                    Caption ="Org"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =600
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label16"
                    Caption ="Spec Name"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1860
                    Top =1680
                    Width =3600
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Spec Element"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =1860
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =87
                    TextFontFamily =18
                    Left =60
                    Top =1380
                    Width =1680
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label28"
                    Caption ="QTS Specs"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6660
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label34"
                    Caption ="Target"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =6660
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1980
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =13140
                    Top =120
                    Width =900
                    FontSize =10
                    FontWeight =700
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="mcrClose"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =7740
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label52"
                    Caption ="Lower"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =1680
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =8820
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label56"
                    Caption ="Upper"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =8820
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =9900
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label60"
                    Caption ="Lower"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =9900
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =10980
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label64"
                    Caption ="Upper"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =10980
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =12060
                    Top =1680
                    Width =2100
                    Height =300
                    FontSize =10
                    Name ="Label68"
                    Caption ="Attribute 1"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =12060
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =14220
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label72"
                    Caption ="Attribute 2"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =14220
                    LayoutCachedTop =1680
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =15300
                    Top =1680
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label76"
                    Caption ="Attribute 3"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =15300
                    LayoutCachedTop =1680
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =7740
                    Top =1440
                    Width =2100
                    Height =300
                    FontSize =10
                    Name ="Label80"
                    Caption ="Specification Limits"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1740
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =9900
                    Top =1440
                    Width =2100
                    Height =300
                    FontSize =10
                    Name ="Label81"
                    Caption ="User Defined Limits"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =9900
                    LayoutCachedTop =1440
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =1740
                End
                Begin Subform
                    Enabled = NotDefault
                    OverlapFlags =85
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =60
                    Top =60
                    Width =3465
                    Height =1260
                    TabIndex =1
                    Name ="frmItemInterfaceQTSHeader"
                    SourceObject ="Form.frmItemInterfaceQTSHeader"

                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =10320
                    Top =120
                    Width =900
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdAdd"
                    Caption ="Add New"
                    OnClick ="mcrQTSNewOpen"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Top =1680
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label88"
                    Caption ="Code"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =5940
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Top =1440
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label90"
                    Caption ="UOM"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =5940
                    LayoutCachedTop =1440
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =1740
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =11400
                    Top =120
                    Width =1560
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdFiberSpecSelect"
                    Caption ="Add Fiber Specs"
                    OnClick ="mcrQTSFiberSelectOpen"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =300
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =60
                    Width =480
                    Height =300
                    ColumnWidth =1185
                    FontSize =10
                    Name ="OrganizationCode"
                    ControlSource ="OrganizationCode"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontFamily =18
                    BackStyle =0
                    Left =600
                    Width =1260
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =1
                    Name ="AssemblyItemNumber"
                    ControlSource ="SpecName"
                    DefaultValue ="=[Forms]![frmItemInterface]![ItemNumber]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6660
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="TargetValue"
                    ControlSource ="TargetValue"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6660
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7740
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="LowerSpecLimit"
                    ControlSource ="LowerSpecLimit"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7740
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =8820
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="UpperSpecLimit"
                    ControlSource ="UpperSpecLimit"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =8820
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =9900
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="LowerUserDefinedLimit"
                    ControlSource ="LowerUserDefinedLimit"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =9900
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =10980
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="UpperUserDefinedLimit"
                    ControlSource ="UpperUserDefinedLimit"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =10980
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =12060
                    Width =2100
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="Attribute1"
                    ControlSource ="Attribute1"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =12060
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =14220
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =9
                    Name ="Attribute2"
                    ControlSource ="Attribute2"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =14220
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =15300
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =10
                    Name ="Attribute3"
                    ControlSource ="Attribute3"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =15300
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =300
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =87
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1860
                    Width =3600
                    Height =300
                    ColumnWidth =2250
                    FontSize =10
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="ElementName"
                    ControlSource ="ElementName"
                    RowSourceType ="Table/Query"
                    RowSource ="qryQTSSpecElementSelect"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1860
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =300
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =5520
                    Width =360
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    Name ="Command82"
                    Caption ="Close"
                    OnClick ="mcrDeleteRecord"
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
                    FontName ="Times New Roman"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =5520
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =300
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5940
                    Width =660
                    Height =300
                    FontSize =10
                    TabIndex =12
                    Name ="UOMCode"
                    ControlSource ="UOMCode"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5940
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            Height =780
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    Left =6060
                    Top =60
                    Width =8460
                    Height =540
                    FontSize =10
                    Name ="Text91"
                    ControlSource ="TargetValue"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    TextFontFamily =18
                    Left =5400
                    Top =60
                    Width =660
                    Height =300
                    FontSize =10
                    Name ="Label92"
                    Caption ="Target"
                    FontName ="Times New Roman"
                End
            End
        End
    End
End

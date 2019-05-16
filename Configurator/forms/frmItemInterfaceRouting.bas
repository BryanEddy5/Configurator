Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6060
    DatasheetFontHeight =10
    ItemSuffix =42
    Left =9090
    Top =5490
    Right =15285
    Bottom =10920
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe84284606cf9e340
    End
    RecordSource ="qryItemInterfaceRoutingView"
    Caption ="frmItemInterface"
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
            Height =1200
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =18
                    Left =900
                    Top =900
                    Width =840
                    Height =300
                    FontSize =10
                    Name ="Label16"
                    Caption ="Item"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =3900
                    Top =660
                    Width =780
                    Height =300
                    FontSize =10
                    Name ="Label27"
                    Caption ="Start Up"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =3900
                    Top =900
                    Width =780
                    Height =300
                    FontSize =10
                    Name ="Label33"
                    Caption ="Scrap"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =95
                    TextFontFamily =18
                    Left =120
                    Top =600
                    Width =1320
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label29"
                    Caption ="Routing"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =3240
                    Top =660
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Oper."
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =2580
                    Top =660
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label23"
                    Caption ="Oper"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =2580
                    Top =900
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label31"
                    Caption ="Seq"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =3240
                    Top =900
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Label35"
                    Caption ="Code"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    TextFontFamily =18
                    Left =360
                    Top =900
                    Width =480
                    Height =300
                    FontSize =10
                    Name ="Label36"
                    Caption ="Org"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =4740
                    Top =660
                    Width =780
                    Height =300
                    FontSize =10
                    Name ="Label39"
                    Caption ="Cable"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =4740
                    Top =900
                    Width =780
                    Height =300
                    FontSize =10
                    Name ="Label41"
                    Caption ="Print"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =4500
                    Top =120
                    Width =1050
                    FontSize =10
                    FontWeight =700
                    Name ="cmdRefresh"
                    Caption ="Refresh"
                    OnClick ="mcrRequery"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =7
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =3900
                    Top =30
                    Width =780
                    Height =300
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =4
                    Name ="StartupScrap"
                    ControlSource ="StartupScrap"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    BackStyle =0
                    Left =900
                    Top =30
                    Width =1620
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =1
                    Name ="AssemblyItemNumber"
                    ControlSource ="AssemblyItemNumber"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =3240
                    Top =30
                    Width =600
                    Height =300
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =3
                    Name ="OperationCode"
                    ControlSource ="OperationCode"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =2580
                    Top =30
                    Width =600
                    Height =300
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =2
                    Name ="OperationSeqNumber"
                    ControlSource ="OperationSeqNum"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =360
                    Top =30
                    Width =480
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    Name ="OrganizationCode"
                    ControlSource ="OrganizationCode"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =4740
                    Top =30
                    Width =780
                    Height =300
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =5
                    Name ="ShortText"
                    ControlSource ="ShortText"
                    FontName ="Times New Roman"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5580
                    Top =15
                    Width =366
                    Height =306
                    TabIndex =6
                    Name ="cmdDelete"
                    Caption ="Command25"
                    OnClick ="mcrDeleteRecord"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End

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
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9180
    DatasheetFontHeight =10
    ItemSuffix =99
    Left =4365
    Top =2610
    Right =15330
    Bottom =7935
    DatasheetGridlinesColor =12632256
    OrderBy ="tblQASpecsFiberDefaults.FiberType"
    RecSrcDt = Begin
        0x12cc837ff2a2e440
    End
    RecordSource ="qryQTSSpecElementsFiberDefaults"
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
            Height =1920
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =480
                    Top =1620
                    Width =2940
                    Height =300
                    FontSize =10
                    Name ="Label17"
                    Caption ="Spec Element"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =3480
                    Top =1620
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label34"
                    Caption ="Target"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =7920
                    Top =480
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
                    Left =4560
                    Top =1620
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label52"
                    Caption ="Lower"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =5640
                    Top =1620
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label56"
                    Caption ="Upper"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =6720
                    Top =1620
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label60"
                    Caption ="Lower"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =18
                    Left =7800
                    Top =1620
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label64"
                    Caption ="Upper"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    TextFontFamily =18
                    Left =4560
                    Top =1380
                    Width =2100
                    Height =300
                    FontSize =10
                    Name ="Label80"
                    Caption ="Specification Limits"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    TextFontFamily =18
                    Left =6720
                    Top =1380
                    Width =2100
                    Height =300
                    FontSize =10
                    Name ="Label81"
                    Caption ="User Defined Limits"
                    FontName ="Times New Roman"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =6840
                    Top =480
                    Width =900
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdApply"
                    Caption ="Apply"
                    OnClick ="mcrQTSFiberSelectApply"
                    FontName ="Times New Roman"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =480
                    Top =540
                    Width =2940
                    Height =300
                    FontSize =10
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboFiberSelect"
                    RowSourceType ="Table/Query"
                    RowSource ="qryQTSSpecElementsFiberSelect"
                    ColumnWidths ="1440;0"
                    AfterUpdate ="mcrRequery"
                    DefaultValue ="=[Forms]![frmMainMenu]![FiberType1]"
                    FontName ="Times New Roman"

                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    TextFontFamily =18
                    Left =4560
                    Top =1080
                    Width =4260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label95"
                    Caption ="DEFAULT VALUES"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =18
                    Left =3480
                    Top =1380
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label96"
                    Caption ="Default"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =480
                    Top =60
                    Width =8280
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label97"
                    Caption ="NOTE:  CHANGES TO DEFAULT VALUES WILL BE AUTOMATICALLY SAVED"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            Height =300
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =3480
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="TargetValue"
                    ControlSource ="DefaultTarget"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =4560
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="LowerSpecLimit"
                    ControlSource ="DefaultLowerLimit"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =5640
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="UpperSpecLimit"
                    ControlSource ="DefaultUpperLimit"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =6720
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="LowerUserDefinedLimit"
                    ControlSource ="DefaultUserLowerLimit"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =7800
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="UpperUserDefinedLimit"
                    ControlSource ="DefaultUserUpperLimit"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =480
                    Width =2940
                    Height =300
                    ColumnWidth =2295
                    FontSize =10
                    Name ="ElementName"
                    ControlSource ="SpecificationElement"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            Height =540
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Left =480
                    Top =120
                    Width =8280
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label98"
                    Caption ="NOTE:  CHANGES TO DEFAULT VALUES WILL BE AUTOMATICALLY SAVED"
                    FontName ="Times New Roman"
                End
            End
        End
    End
End

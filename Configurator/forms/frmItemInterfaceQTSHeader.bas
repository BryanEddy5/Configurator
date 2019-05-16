Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ScrollBars =0
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3180
    DatasheetFontHeight =10
    ItemSuffix =86
    Left =4365
    Top =1485
    Right =8985
    Bottom =4710
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc25fe191124ce440
    End
    RecordSource ="qryItemInterfaceQTSViewHeader"
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
            Height =660
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Top =360
                    Width =480
                    Height =300
                    FontSize =10
                    Name ="Label15"
                    Caption ="Org"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1920
                    Top =360
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
                    Left =600
                    Top =360
                    Width =1260
                    Height =300
                    FontSize =10
                    Name ="Label83"
                    Caption ="Item Number"
                    FontName ="Times New Roman"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    Width =2280
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label85"
                    Caption ="Oracle Specification"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin Section
            Height =360
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
                    Top =30
                    Width =480
                    Height =300
                    ColumnWidth =1185
                    FontSize =10
                    Name ="OrganizationCode"
                    ControlSource ="OrganizationCode"
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
                    Left =1920
                    Top =30
                    Width =1260
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =1
                    Name ="AssemblyItemNumber"
                    ControlSource ="SpecName"
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
                    Left =600
                    Top =30
                    Width =1260
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Item"
                    ControlSource ="Item"
                    FontName ="Times New Roman"

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

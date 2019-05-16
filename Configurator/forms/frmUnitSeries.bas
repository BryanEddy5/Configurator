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
    TabularFamily =22
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =5235
    Top =3480
    Right =11340
    Bottom =8910
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7fc4c363dd92e440
    End
    RecordSource ="qryUnitSeriesExisting"
    Caption ="frmUnitSeries"
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
            Height =405
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1230
                    Top =60
                    Width =2025
                    Height =240
                    FontSize =10
                    Name ="Exists_Label"
                    Caption ="Existing Unit Series"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =480
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =960
                    Top =60
                    Width =2700
                    Height =255
                    ColumnWidth =900
                    FontSize =10
                    Name ="Existing"
                    ControlSource ="Unit Series"
                    FontName ="Tahoma"

                End
            End
        End
        Begin FormFooter
            Height =1980
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =1260
                    Width =1680
                    Height =405
                    TabIndex =1
                    Name ="btnfrmUnitSeries"
                    Caption ="Save / Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Top =540
                    Width =3000
                    FontSize =10
                    FontWeight =600
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="NewUnit"
                    RowSourceType ="Table/Query"
                    RowSource ="Unit Series"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =180
                            Width =2100
                            Height =240
                            FontSize =10
                            Name ="Label4"
                            Caption ="Enter New Unit Series:"
                            FontName ="Tahoma"
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
Option Explicit
Option Compare Database

Private Sub Save___Close_Click()
    On Error GoTo Err_Save___Close_Click

    Dim stDocName As String

    stDocName = "mcrPrintType"
    DoCmd.RunMacro stDocName

Exit_Save___Close_Click:
    Exit Sub

Err_Save___Close_Click:
    MsgBox Err.Description
    Resume Exit_Save___Close_Click
    
End Sub

Private Sub btnfrmUnitSeries_Click()
    DoCmd.OpenForm "frmUnitSeriesExisting"

    If IsNull(Me.NewUnit) = True Or Forms!frmUnitSeriesExisting!Existing = -1 Then
        MsgBox ("Plese create new Unit Series")
    Else
        Forms!frmCreateItem!UnitSeries = Me.NewUnit
        DoCmd.Close acForm, "frmUnitSeries"

    End If

    DoCmd.Close acForm, "frmUnitSeriesExisting"
    


End Sub

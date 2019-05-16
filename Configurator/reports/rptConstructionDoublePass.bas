Version =20
VersionRequired =20
Begin Report
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9180
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =8940
    Top =2340
    DatasheetGridlinesColor =12632256
    Filter ="Base ='BM006-2.5'"
    RecSrcDt = Begin
        0x4e009d8acefee440
    End
    RecordSource ="SELECT tblCableConstructionReferences.Base, tblCableConstructionReferences.DateC"
        "reated, tblCableConstructionReferences.CreatedBy, tblCableConstructionReferences"
        ".RevisionDate, tblCableConstructionReferences.RevisedBy, tblCableConstructionRef"
        "erences.DesignCode, tblCableConstructionReferences.DesignExtension, tblCableCons"
        "tructionReferences.ConstructionDescription, tblCableConstructionReferences.BaseI"
        "D, tblCableConstructionReferences.NumSubFillers, tblCableConstructionReferences."
        "Active, tblCableConstructionReferences.PrefixID, tblCableConstructionReferences."
        "ProductID, tblCableConstructionReferences.SetupID, tblCableConstructionReference"
        "s.FamilyID, tblCableConstructionReferences.UnitIDTypeNumber, tblCableConstructio"
        "nReferences.InactiveReason, tblCableConstructionReferences.InactiveReason, tblCa"
        "bleConstructionReferences.NumSubPositions, tblCableConstructionReferences.TBType"
        " FROM tblCableConstructionReferences; "
    Caption ="Cable Construction"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000dc230000e04c000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =3
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =90.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =19680
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    Width =9180
                    Height =19440
                    BorderColor =15916467
                    Name ="cConstruction"
                    SourceObject ="Form.frmConstructionDoublePass"
                    LinkChildFields ="Base"
                    LinkMasterFields ="Base"
                    GridlineColor =10921638

                    LayoutCachedWidth =9180
                    LayoutCachedHeight =19440
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

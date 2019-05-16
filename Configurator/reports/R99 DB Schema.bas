Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10490
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =945
    Top =495
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb2d466c4b60ee240
    End
    RecordSource ="~~my test~~"
    Caption ="DB Schema"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa802000089050000a80200008905000000000000fa2800005502000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            Width =850
            Height =850
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="TableName"
        End
        Begin BreakLevel
            ControlSource ="FieldSeq"
        End
        Begin BreakLevel
            ControlSource ="FieldName"
        End
        Begin PageHeader
            Height =1320
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =57
                    Top =822
                    Width =1311
                    Height =315
                    Name ="TableName_Label"
                    Caption ="TableName"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1425
                    Top =822
                    Width =2109
                    Height =315
                    Name ="FieldName_Label"
                    Caption ="Field"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3591
                    Top =822
                    Width =1083
                    Height =315
                    Name ="FieldType_Label"
                    Caption ="Type"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =4731
                    Top =822
                    Width =567
                    Height =315
                    Name ="FieldSize_Label"
                    Caption ="Size"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =5387
                    Top =822
                    Width =2121
                    Height =315
                    Name ="FieldPurpose_Label"
                    Caption ="Purpose"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =765
                    Width =9075
                    Name ="Line14"
                End
                Begin Line
                    BorderWidth =2
                    Top =1194
                    Width =9075
                    Name ="Line15"
                End
                Begin Label
                    BackStyle =1
                    Left =57
                    Top =60
                    Width =2040
                    Height =525
                    FontSize =20
                    Name ="Label16"
                    Caption ="DB Schema"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9075
                    BorderColor =12632256
                    Name ="Line17"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =3
                    Left =3384
                    Top =60
                    Width =5688
                    Height =525
                    FontSize =10
                    Name ="DBName_lab"
                    Caption ="Database name"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =390
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =57
                    Width =5916
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="TableName"
                    ControlSource ="TableName"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =675
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    Left =900
                    Top =225
                    Width =2634
                    Height =225
                    Name ="FieldName"
                    ControlSource ="FieldName"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    ColumnCount =2
                    Left =3591
                    Top =225
                    Width =1083
                    Height =225
                    TabIndex =1
                    Name ="FieldType"
                    ControlSource ="FieldType"
                    RowSourceType ="Value List"
                    RowSource ="1;dbBoolean;2;dbByte;3;dbInteger;4;dbLong;5;dbCurrency;6;dbSingle;7;dbDouble;8;d"
                        "bDate;9;dbBinary;10;dbText;11;dbLongBinary;12;dbMemo;15;dbGUID;16;dbBigInt;17;db"
                        "VarBinary;18;dbChar;19;dbNumeric;20;dbDecimal;21;dbFloat;22;dbTime;23;dbTimestam"
                        "p"
                    ColumnWidths ="0"

                End
                Begin TextBox
                    Left =4731
                    Top =225
                    Width =567
                    Height =225
                    TabIndex =2
                    Name ="FieldSize"
                    ControlSource ="FieldSize"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    Left =5387
                    Top =225
                    Width =5103
                    Height =225
                    ColumnWidth =6675
                    TabIndex =3
                    Name ="FieldPurpose"
                    ControlSource ="FieldPurpose"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    Left =5387
                    Top =450
                    Width =5103
                    Height =225
                    TabIndex =4
                    Name ="Text18"
                    ControlSource ="FieldComment"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    Left =900
                    Width =8178
                    Height =225
                    TabIndex =5
                    Name ="TableInfo"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =195
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =528
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =57
                    Top =228
                    Width =4424
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text11"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =4594
                    Top =228
                    Width =4424
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text12"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

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


Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

'   concatenate Table-level stuff
Static currTable$, tabInfo$
    If currTable <> Nz(Me![TableName]) Then
        currTable = Nz(Me![TableName])
        tabInfo = ""
        End If
    If Left$(currTable, 10) = " Relation:" Then
        Me!FieldType.Visible = False
        Me!FieldSize.Visible = False
        Me!FieldName.Width = 7.741 * fig_TwipsPerCm
        Else
        Me!FieldType.Visible = True
        Me!FieldSize.Visible = True
        Me!FieldName.Width = 4.646 * fig_TwipsPerCm
        End If
    If (Left$(Me![FieldName], 5) = "Prop ") Or _
       (Left$(Me![FieldName], 6) = "Index ") Then
        tabInfo = tabInfo & Me![FieldName] & "  " & Me![FieldPurpose] & vbCrLf
        Me.NextRecord = True
        Me.MoveLayout = False
        Me.PrintSection = False
        Else
        Me!TableInfo = tabInfo
        tabInfo = ""
        End If
    End Sub

Private Sub Report_Open(Cancel As Integer)
    
    Dim col_oa As Collection
    Dim sql$, tbl$
    Dim schema_tableName$, schema_workDbName$, schema_DbName$
    
    Set col_oa = glb_coll(1)
    glb_coll.Remove 1
    schema_tableName$ = col_oa(1)
    schema_workDbName$ = col_oa(2)
    schema_DbName$ = col_oa(3)
    Set col_oa = Nothing
    tbl$ = "[" & schema_tableName$ & "]"
    sql$ = ""
    sql$ = sql$ & " SELECT " & tbl$ & ".* "
    sql$ = sql$ & " FROM " & tbl$
    If schema_workDbName$ <> "" Then
        sql$ = sql$ & " IN """ & schema_workDbName$ & """"
        End If
    Me.RecordSource = sql$
    
    Me![DBName_lab].Caption = schema_DbName$
    End Sub

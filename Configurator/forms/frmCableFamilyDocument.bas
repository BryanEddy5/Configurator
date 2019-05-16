Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6480
    DatasheetFontHeight =10
    ItemSuffix =79
    Left =540
    Top =3165
    Right =7275
    Bottom =8670
    DatasheetGridlinesColor =12632256
    Filter ="((tblBurnListing.Base Like \"ar*\"))"
    OrderBy ="tblBurnListing.Base"
    RecSrcDt = Begin
        0x6f26435ccfb8e440
    End
    RecordSource ="tblBurnListing"
    Caption ="Item Search"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =4020
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =3480
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Document Name"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3480
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =3780
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =420
                    Top =60
                    Width =4620
                    Height =780
                    Name ="Box35"
                    LayoutCachedLeft =420
                    LayoutCachedTop =60
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =840
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1860
                    Width =6000
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1860
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =540
                    Top =180
                    Width =4215
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Add Cable Family Doc"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =540
                    LayoutCachedTop =180
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =675
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =3480
                    Width =2895
                    Height =300
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =500
                    Name ="CableDocument"
                    Format ="@;\" <enter document name>\""

                    LayoutCachedLeft =2280
                    LayoutCachedTop =3480
                    LayoutCachedWidth =5175
                    LayoutCachedHeight =3780
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =2040
                    Width =3555
                    Height =900
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =2
                    Name ="FileList"
                    RowSourceType ="Value List"
                    RowSource ="C:\\Users\\eddybc\\Desktop\\N2015-1061 Release Review"
                    FontName ="Tahoma"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =2280
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5835
                    LayoutCachedHeight =2940
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =2040
                    Width =1086
                    Height =606
                    TabIndex =1
                    Name ="btnAttachSpec"
                    Caption ="Attach Doc"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"

                    LayoutCachedLeft =180
                    LayoutCachedTop =2040
                    LayoutCachedWidth =1266
                    LayoutCachedHeight =2646
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5220
                    Top =3480
                    Width =645
                    Height =300
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="FileType"

                    LayoutCachedLeft =5220
                    LayoutCachedTop =3480
                    LayoutCachedWidth =5865
                    LayoutCachedHeight =3780
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =3120
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label73"
                    Caption ="Cable Family"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3120
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =3420
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =10080
                    Left =2280
                    Top =3120
                    Width =3555
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"300\""
                    Name ="CableFamily"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblCableFamily.FamilyID, tblCableFamily.FamilyDescription AS Family, tblC"
                        "ableFamily.NewProductRelease AS NPR, tblCableFamily.EngineeringChange AS EC, tbl"
                        "CableFamily.NewProductCode AS NPC FROM tblCableFamily ORDER BY tblCableFamily.Fa"
                        "milyID DESC; "
                    ColumnWidths ="0;3600;2160;1440;1440"
                    Format ="@;\" <select cable family>\""
                    LayoutCachedLeft =2280
                    LayoutCachedTop =3120
                    LayoutCachedWidth =5835
                    LayoutCachedHeight =3375
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =240
                    Top =1200
                    Width =960
                    Height =480
                    TabIndex =5
                    Name ="btnAddDocs"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000024000000240000000100180000000000300f0000c40e0000c40e0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffff3f3f3cccdcdc9cacac9ca ,
                        0xcac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9 ,
                        0xcacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9cacac9caca ,
                        0xcacacae7e7e7ffffffffffffffffffffffffffffffffffffffffffffffffc3c3 ,
                        0xc35a5b5b56575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x565757565757565757565757565757575858adadadffffffffffffffffffffff ,
                        0xfffffffffffffff0f1f15b5c5c56575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57d2d2d2ffffffffffffffffffffffffffffffc4c4c456575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x575657575657575657575657577f8080ffffffffffffffffffffffffffffffb1 ,
                        0xb1b1565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757565757565757565757565757565757565757f4f4f4ff ,
                        0xffffffffffffffffffffffabacac565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757d8d8d8ffffffffffffffffffffffffabacac565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757babbbbffffffffffffffffff ,
                        0xffffffabacac5657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x9c9c9cffffffffffffffffffffffffabacac5657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657577f8080ffffffffffffffffffffffffabacac5657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x565757565757565757565757565757565757565757565757636363ffffffffff ,
                        0xffffffffffffffabacac56575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757f1f1f1ffffffffffffffffffabacac56575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757565757565757565757d1d1d1ffffffffffffffffffab ,
                        0xacac565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757565757565757565757565757565757565757565757b2 ,
                        0xb3b3ffffffffffffffffffabacac565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757969696ffffffffffffffffffabacac565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757777878ffffffffffff ,
                        0xffffffabacac5657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575e5f5ffdfdfdffffffffffffabacac5657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x565757565757565757565757565757565757eaeaeaffffffffffffabacac5657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x565757565757565757565757565757565757565757565757565757565757c9c9 ,
                        0xc9ffffffffffffabacac56575756575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757b0b1b1ffffffffffffabacac5657578d8e8e595a5a56 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x57565757565757565757565757565757565757565757bababaffffffffffffaa ,
                        0xabab565757b6b7b7d5d5d5727373555656565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575770 ,
                        0x7070f4f4f4ffffffffffffa9a9a9565757b6b7b7ffffffffffffe7e7e7dbdbdb ,
                        0xdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdb ,
                        0xdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdbdb ,
                        0xdbdbdbdbdbd7d7d7cccdcdfcfcfcffffffffffffffffffa9a9a9565757afb0b0 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffb4b4b4e1e1e1ffffffffffffffffff ,
                        0xffffffa9a9a95657576b6b6beaebebffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffff8f8f8afafaf585959 ,
                        0xe1e1e1ffffffffffffffffffffffffb4b4b4565757565757595a5a5f60605f60 ,
                        0x605f60605f60605f60605f60605f60605f60605f60605f60605f60605f60605f ,
                        0x60605f60605f6060cfcfcfffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffa9aaaa565757565757e5e5e5ffffffffffffffffffffffffdfdfdf5657 ,
                        0x5756575756575756575756575756575756575756575756575756575756575756 ,
                        0x57575657575657575657575657575657575657576c6d6dc5c5c5d1d1d1d1d1d1 ,
                        0xd1d1d1d1d1d1d1d1d1d1d1d1b8b9b95d5e5e565757696a6affffffffffffffff ,
                        0xffffffffffffffffffffa2a3a356575756575756575756575756575756575756 ,
                        0x5757565757565757565757565757565757565757565757565757565757565757 ,
                        0x5657575657575657575657575657575657575657575657575657575657575b5c ,
                        0x5cd4d4d4ffffffffffffffffffffffffffffffffffffffffffcecece9d9d9d91 ,
                        0x9292919292919292919292919292919292919292919292919292919292919292 ,
                        0x9192929192929192929192926465655657575657575657575657575657575657 ,
                        0x575657575657577c7d7de5e6e6ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffaaabab5657575657 ,
                        0x57565757565757565757565757565757565757cecfcfffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffdfdfdd7d7d7c9cacac9cacac9cacac9cacac9cacac9cacae1e1e1ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add NPR Documents"
                    Picture ="NPR.bmp"

                    LayoutCachedLeft =240
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =1680
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =0
            BackColor =-2147483633
            Name ="Detail"
        End
        Begin FormFooter
            Height =1500
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =660
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =4260
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="btnSubmit"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4260
                    LayoutCachedTop =660
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =1020
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
Private WithEvents m_frmCableFamily As Form_frmCableFamily
Attribute m_frmCableFamily.VB_VarHelpID = -1

Private Sub btnAddDocs_Click()

If IsNull(Me.CableFamily) Then
    Exit Sub
End If


DoCmd.OpenForm "frmCableFamilyViewDocument"

ShowCableFamilyDocuments (Me.CableFamily)

Forms.frmCableFamilyViewDocument.bxDocList.Selected(0) = True
End Sub

Private Sub btnAttachSpec_Click()
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
   Dim fileName As String
   Dim FileType As String
   Dim filePath As String
 
 On Error GoTo ErrorHandler:
   ' Clear listbox contents.
   Me.FileList.rowSource = ""
 
   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
 
   With fDialog
 
      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = True
             
      ' Set the title of the dialog box.
      .Title = "Please select one or more files"
 
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "All Files", "*.*"
      .Filters.Add "Access Databases", "*.MDB"
      .Filters.Add "Access Projects", "*.ADP"
      
 
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
 
         'Loop through each file selected and add it to our list box.
         
         For Each varFile In .SelectedItems
            Me.FileList.AddItem varFile
         Next
        
        DoCmd.SetWarnings False
        'Create document name from document path
        filePath = FileList.ItemData(0)
        FileType = Mid(filePath, InStrRev(filePath, "."), Len(filePath) - (InStrRev(filePath, ".") - 1))
        fileName = InStrRev(filePath, "\")
        Me.CableDocument = Mid(filePath, fileName + 1, Len(filePath) - fileName - Len(FileType))
        Me.FileType = FileType
        
        
            
        DoCmd.SetWarnings True
      Else
         'MsgBox "You clicked Cancel in the file dialog box."
      End If
   End With
Exit Sub
   
ErrorHandler:
MsgBox "Command has produced an error: " & Err.Number & " " & Err.Description
   
   
End Sub

Private Sub btnClose_Click()
DoCmd.Close
End Sub


Private Sub btnSubmit_Click()

    Dim strDocument As String
    Dim intFamilyID As Integer
    Dim strFilePath As String
    Dim NewPath As String
    Dim SaveDocSuccess As Boolean
    Dim strDocumentName As String
    Dim FamilyID%, i%
    Dim doc$, newDoc$, path$, Source$, Destination$
    On Error GoTo ErrorHandler:
    
    If MsgBox("Are you ready to submit?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If IsNull(Me.FileList.ItemData(0)) = True Then
        If MsgBox("Do you want to attach a Spec Sheet?", vbYesNo) = vbYes Then
            Call btnAttachSpec_Click
            Exit Sub
        End If
    End If
    
    FamilyID = Me.CableFamily
'    strDocument = Me.CableDocument
'    strFilePath = Me.FileList.ItemData(0)
    path = Core_PathwayManager.GetPathway(NPR)
'    strDocumentName = intFamilyID & "." & strDocument
'
    For i = 0 To Me.FileList.ListCount - 1
        
        'Modify the document to have the cablefamilyid as a prefix
        Source = Me.FileList.ItemData(i)
        doc = GetFileName(Source)
        Destination = path & FamilyID & "." & doc
        Debug.Print "Source: " & Source & vbCrLf & "Destination: " & Destination
        
        If CommonFunctions.FileExists(Destination) Then
            If MsgBox("File: " & doc & " already exists.  Do you want to overwrite it?", vbYesNo) = vbYes Then
                 FileCopy Source, Destination
            End If
        Else
             FileCopy Source, Destination
             Call SaveDocumentTable(FamilyID, doc)
        End If
    
    Next
    SaveDocSuccess = True
    

    If SaveDocSuccess = True Then
        
        Beep
        MsgBox ("Load is Complete")
    Else
        MsgBox "Error occurred while saving document. Ensure document is not open.", vbCritical
    End If

    
    
    
ErrorExit:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description & " " & Err.Number
    Resume ErrorExit


End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err_cmdDelete_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
End Sub


Private Sub Form_Load()
Me.FileList.rowSource = ""

Set m_frmCableFamily = Form_frmCableFamily

End Sub


Function SaveDocumentTable(intFamilyID As Integer, strDocument As String)

Dim strSQL As String

strSQL = "INSERT INTO tblCableFamilyDocs (FamilyID, Document) VALUES  ('" & intFamilyID & "'  , '" & strDocument & "') "

DoCmd.RunSQL strSQL

End Function


Private Sub m_frmCableFamily_Selected(value As String)

Me.CableFamily = value

End Sub

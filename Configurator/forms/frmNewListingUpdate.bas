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
    Width =5700
    DatasheetFontHeight =10
    ItemSuffix =79
    Left =5235
    Top =3195
    Right =11190
    Bottom =10500
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
                    OverlapFlags =85
                    Left =180
                    Top =3000
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Document"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3000
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =3300
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
                    Top =1380
                    Width =5160
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1380
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1380
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
                    Caption ="Add / Update Listing"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =540
                    LayoutCachedTop =180
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =675
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =1560
                    Width =2940
                    Height =900
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =1
                    Name ="FileList"
                    RowSourceType ="Value List"
                    FontName ="Tahoma"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =2280
                    LayoutCachedTop =1560
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2460
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
                    Top =1560
                    Width =1086
                    Height =606
                    Name ="btnAttachSpec"
                    Caption ="Attach Spec"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1560
                    LayoutCachedWidth =1266
                    LayoutCachedHeight =2166
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4560
                    Top =3000
                    Width =645
                    Height =300
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =500
                    TabIndex =4
                    Name ="FileType"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =3000
                    LayoutCachedWidth =5205
                    LayoutCachedHeight =3300
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =2640
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label73"
                    Caption ="Listing Company"
                    LayoutCachedLeft =180
                    LayoutCachedTop =2640
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =2940
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =2640
                    Width =2235
                    Height =255
                    ColumnOrder =2
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="ListingCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="tblBurnListingCompany"
                    AfterUpdate ="[Event Procedure]"
                    Format ="@;\" <pick company>\""

                    LayoutCachedLeft =2280
                    LayoutCachedTop =2640
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =2895
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =3420
                    Width =1695
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label75"
                    Caption ="Rename Doc To"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3420
                    LayoutCachedWidth =1875
                    LayoutCachedHeight =3720
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =4620
                    Top =3510
                    Width =240
                    ColumnOrder =1
                    TabIndex =5
                    BorderColor =15916467
                    Name ="chkReplace"
                    GridlineColor =10921638

                    LayoutCachedLeft =4620
                    LayoutCachedTop =3510
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =3750
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =119
                            Left =4850
                            Top =3480
                            Width =690
                            Height =240
                            Name ="Label78"
                            Caption ="Replace"
                            LayoutCachedLeft =4850
                            LayoutCachedTop =3480
                            LayoutCachedWidth =5540
                            LayoutCachedHeight =3720
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =4320
                    Left =2280
                    Top =3000
                    Width =2235
                    Height =255
                    FontSize =9
                    FontWeight =500
                    TabIndex =6
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Listing"
                    RowSourceType ="Table/Query"
                    RowSource ="qryBurnListingDocumentUPdate"
                    Format ="@;\" <select document>\""

                    LayoutCachedLeft =2280
                    LayoutCachedTop =3000
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =3255
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =3420
                    Width =2235
                    Height =255
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="ReplaceDocument"
                    Format ="@;\" <rename doc>\""

                    LayoutCachedLeft =2280
                    LayoutCachedTop =3420
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =3675
                End
            End
        End
        Begin Section
            Height =1800
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

Private Sub btnAttachSpec_Click()
    Dim fDialog As Office.FileDialog
    Dim varFile As Variant
    Dim fileName As String
    Dim FileType As String
    Dim filePath As String
 
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
        
            'Create document name from document path
            filePath = FileList.ItemData(0)
            FileType = Mid(filePath, InStrRev(filePath, "."), Len(filePath) - (InStrRev(filePath, ".") - 1))
            fileName = InStrRev(filePath, "\")
            'Me.Listing = Mid(Filepath, FileName + 1, Len(Filepath) - FileName - Len(FileType))
            Me.FileType = FileType

        Else
            'MsgBox "You clicked Cancel in the file dialog box."
        End If
    End With
   
   
End Sub

Private Sub btnClose_Click()
    DoCmd.Close
End Sub

Private Sub btnSubmit_Click()
    DoCmd.SetWarnings False
    Dim filePath$
    On Error GoTo ErrorHandler:

    If MsgBox("Are you ready to submit?", vbYesNo) = vbNo Then
        Exit Sub
    End If

    If IsNull(Me.FileList.ItemData(0)) = True Then
        If MsgBox("Do you want to attach a Spec Sheet?", vbYesNo) = vbYes Then
            Call btnAttachSpec_Click
            Exit Sub
        Else
            Exit Sub
            MsgBox ("Update is Complete")
        End If
    End If
    
    If IsNull(Me.ReplaceDocument) = False Then
        'DoCmd.OpenQuery "qryUpdateListingDocument"
        DoCmd.RunSQL "UPDATE tblBurnListingDocument SET ListingDocument = '" & Me.ReplaceDocument & "' WHERE ListingDocument = '" & Me.Listing & "'"
        'DoCmd.RunSQL "UPDATE tblBurnListingConstruction SET ListingDocument = '" & Me.Listing & "' WHERE ListingDocument = '" & Me.ReplaceDocument & "'"
    Else
        DoCmd.RunSQL "INSERT INTO tblBurnListingDocument (ListingCompany, ListingDocument) VALUES ( '" & Me.ListingCompany & "' , '" & Me.Listing & "')"
    End If
    
    
    
    filePath = Me.FileList.ItemData(0)
    
        
    
    UpdateBurnListingFile Me.Listing, filePath, Me.FileType, Me.ListingCompany
    
    
    


    Beep
    MsgBox ("Load is Complete")

ErrorExit:
    Exit Sub

ErrorHandler:
    ''MsgBox Err.Description & " " & Err.Number
    Resume Next


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

End Sub

Private Sub ListingCompany_AfterUpdate()
    Me.ReplaceDocument = Null
    Me.ReplaceDocument.Requery
    Me.Listing = Null
    Me.Listing.Requery
End Sub

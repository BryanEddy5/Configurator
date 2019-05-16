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
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5700
    DatasheetFontHeight =10
    ItemSuffix =72
    Left =11265
    Top =1755
    Right =17220
    Bottom =11040
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0e1aacebf4b1e440
    End
    RecordSource ="tblCustomerSpecsHold"
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
                    Top =2700
                    Width =1560
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Document"
                    LayoutCachedLeft =180
                    LayoutCachedTop =2700
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =3000
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
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =3120
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Customer"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3120
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =3420
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
                    Caption ="Add New Customer Spec"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =540
                    LayoutCachedTop =180
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =675
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =3540
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label68"
                    Caption ="Rev"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3540
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =3840
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =2700
                    Width =2940
                    Height =300
                    ColumnOrder =2
                    FontSize =9
                    FontWeight =500
                    Name ="Spec"
                    Format ="@;\" <spec name>\""

                    LayoutCachedLeft =2280
                    LayoutCachedTop =2700
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =3000
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =3540
                    Width =2940
                    Height =300
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="CustomerRev"
                    Format ="@;\" <enter spec rev>\""

                    LayoutCachedLeft =2280
                    LayoutCachedTop =3540
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =3840
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =1620
                    Width =2940
                    Height =900
                    ColumnOrder =1
                    FontSize =8
                    TabIndex =4
                    Name ="FileList"
                    RowSourceType ="Value List"
                    FontName ="Tahoma"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =2280
                    LayoutCachedTop =1620
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2520
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
                    Top =1680
                    Width =1086
                    Height =606
                    TabIndex =3
                    Name ="btnAttachSpec"
                    Caption ="Attach Spec"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1266
                    LayoutCachedHeight =2286
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =3120
                    Width =2940
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Customer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT vCustomers.Customer FROM vCustomers; "
                    Format ="@;\" <pick customer>\""

                    LayoutCachedLeft =2280
                    LayoutCachedTop =3120
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =3420
                End
            End
        End
        Begin Section
            Height =426
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =60
                    Width =366
                    Height =366
                    Name ="cmdDelete"
                    Caption ="Command12"
                    OnClick ="[Event Procedure]"
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
                    ControlTipText ="Delete Record"

                    LayoutCachedLeft =4740
                    LayoutCachedTop =60
                    LayoutCachedWidth =5106
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Customer Part #"
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2280
                    Top =60
                    Width =2280
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="Text71"
                    ControlSource ="CustomerPartNumber"
                    Format ="@;\" <enter PID>\""

                    LayoutCachedLeft =2280
                    LayoutCachedTop =60
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =360
                End
            End
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

                    LayoutCachedLeft =300
                    LayoutCachedTop =660
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =1020
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
    Dim filePath As String
    Dim FileType As String
    Dim fileName As String
 
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
         
            filePath = FileList.ItemData(0)
            FileType = Mid(filePath, InStrRev(filePath, "."), Len(filePath) - (InStrRev(filePath, ".") - 1))
            fileName = InStrRev(filePath, "\")
            Me.Spec = Mid(filePath, fileName + 1, Len(filePath) - fileName - Len(FileType))
            'Me.FileType = FileType
            Me.CustomerRev = Null
 
        Else
            'MsgBox "You clicked Cancel in the file dialog box."
        End If
    End With
End Sub

Private Sub btnClose_Click()
    DoCmd.Close
End Sub

Private Sub btnSubmit_Click()

    On Error GoTo ErrorHandler:

    Dim CustomerSpec As String
    Dim SpecRev As String
    Dim Customer As String
    Dim strCustomerPartNumber As String
    Dim SpecPath As String
    
    CustomerSpec = Me.Spec
    SpecRev = Me.CustomerRev
    Customer = Me.Customer
    
    Me.Refresh
    Me.Requery
    
    If Len(Me.CustomerRev) > 2 Then
        If MsgBox("Customer Rev is longer than 2. Is this accurate?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    If IsNull(Me.FileList.ItemData(0)) = True Then
        If MsgBox("Do you want to attach a Spec Sheet?", vbYesNo) = vbYes Then
            Call btnAttachSpec_Click
            Exit Sub
        Else
            GoTo LoadSpectoTable
        End If
    End If
    

    SpecPath = Me.FileList.ItemData(0)
    
        
    
    Call SaveCustomerSpec(CustomerSpec, SpecRev, Customer, SpecPath)
    
    If IsNull(Me.CustomerPartNumber) = True Then
        Beep
        MsgBox ("Load is Complete")
        Exit Sub
    End If


LoadSpectoTable:
    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.connection
    
    strCustomerPartNumber = Me.CustomerPartNumber
    
    Set cnn = CurrentProject.connection
    rst.Open "tblCustomerSpecsHold", cnn, adOpenKeyset, adLockOptimistic
    rst.MoveFirst
    Do Until rst.EOF
        strCustomerPartNumber = rst!CustomerPartNumber
        Call SaveSpectoTable(Customer, CustomerSpec, strCustomerPartNumber)
        rst.MoveNext
        
    Loop
    
    cnn.Close
    Set cnn = Nothing
    Set rst = Nothing

    Beep
    MsgBox ("Load is Complete")

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

    DoCmd.RunSQL "DELETE * FROM tblCustomerSpecsHold"
    Me.Refresh
    Me.Requery

    Me.FileList.rowSource = ""

End Sub

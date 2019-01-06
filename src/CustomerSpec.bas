Attribute VB_Name = "CustomerSpec"
Option Compare Database
Option Explicit
Private Const Path As String = "\\naa.fujikurausa.com\dfs\spb\Teams\Cable_Tech\Customers\"



Function SaveCustomerSpec(CustomerSpec As String, SpecRev As String, Customer As String, SpecPath As String)


    Dim FileType As String
    Dim NewPath As String
    Dim NewFile As String
    
    'SpecPath = Forms.frmCreateItem.FileList.ItemData(0)
    FileType = Mid(SpecPath, InStrRev(SpecPath, "."), Len(SpecPath) - (InStrRev(SpecPath, ".") - 1))
    NewFile = CustomerSpec & "." & SpecRev & FileType
    NewPath = Path & Customer & "\"
    
    
    'Finds if the path exists
    If Not Dir(NewPath) = vbNullString Then
          'Find if the file does not does not exist
        If Dir(NewPath & NewFile) = vbNullString Then
            'Creates new file if one does not exist
            FileCopy SpecPath, NewPath & NewFile
        Else
            'Asks user if they would like to write over the file
            'strResponse = MsgBox("File already exists. Write over file?", vbYesNo + vbCritical, "Override File")
            'If strResponse = vbYes Then
                'Writes over existing file
                'FileCopy Path, NewPath & NewFile
            'ElseIf strResponse = vbNo Then
               'Exit Function
            'End If
        End If
    Else
        'Creates new path if one does not exist
        MkDir (NewPath)
        'Creates new file in the file path
        FileCopy SpecPath, NewPath & NewFile
    End If
    

            

End Function

Function UpdateCustomerSpecName(Customer As String, Document As String, newDocument As String) As Boolean
    Dim fld As Object
    Dim fl As Object 'File
    Dim FSO As Object
    
    'On Error GoTo Error_Resume
    Set FSO = CreateObject("scripting.FileSystemObject") ' late binding
    
    Set fld = FSO.GetFolder(Path & Customer)
    
    For Each fl In fld.Files
        If fl.Name Like Document & "*" Then


            fl.Name = Replace(fl.Name, Document, newDocument)
            'Debug.Print fl.Name
            UpdateCustomerSpecName = True
            
        End If
    Next
    
'Error_Resume:
'    Resume Next


End Function


Function OpenCustomerSpec(strPath As String, File As String)


    Dim Shex As Object
    Dim sNewPath As String
    Set Shex = CreateObject("Shell.Application")
    Dim FileNumber As Integer
    Dim hold() As String
    
    FileNumber = 1
    hold = Split(File, ".")
    

    sNewPath = Environ("USERPROFILE") & "\Downloads\" & File
    If Dir(sNewPath) = vbNullString Then
        FileCopy strPath, sNewPath
        
    Else
        While Not Dir(sNewPath) = vbNullString
            If UBound(hold) < 2 Then
                sNewPath = Environ("USERPROFILE") & "\Downloads\" & hold(0) & "(" & FileNumber & ")." & hold(1)
            Else
                sNewPath = Environ("USERPROFILE") & "\Downloads\" & hold(0) & "." & hold(1) & "(" & FileNumber & ")." & hold(2)
            End If
            FileNumber = 1 + FileNumber
        Wend
        
        FileCopy strPath, sNewPath
    End If
    
    
    
    Shex.Open (sNewPath)
    





End Function



Function ShowCustomerSpec(CustomerPartNumber As String, Customer As String)

Dim strPath As String
Dim File As Variant
Dim CustomerSpec As String
Dim cnn As ADODB.Connection
Dim rst As New ADODB.Recordset
Set cnn = CurrentProject.Connection

DoCmd.OpenForm "frmCustomerSpecList"

Forms.frmCustomerSpecList.bxSpecList.rowSource = ""


rst.Open "SELECT * FROM tblCustomerSpecs WHERE CustomerPartNumber = '" & CustomerPartNumber & "' AND Customer = '" & Customer & "'", cnn, adOpenKeyset, adLockOptimistic



'rst.filter = "CustomerPartNumber = '" & CustomerPartNumber & "'"
rst.MoveFirst

Do Until rst.EOF

CustomerSpec = rst!CustomerSpec

strPath = Path & Customer & "\" & CustomerSpec & "*"
'strPathSpec = Path & Customer & "\"

File = Dir(strPath)
   While (File <> "")
      'If InStr(file, CustomerSpec) > 0 Then
        'strFileName = Path & "\" & File & vbCrLf & strFileName
        'strFile = strPathSpec & File
        'strFile = File
        Forms.frmCustomerSpecList.bxSpecList.AddItem File
         'MsgBox "found " & vbCrLf & strFileName
         'Exit Function
      'End If
     File = Dir
  Wend
  
rst.MoveNext
  
Loop

Forms.frmCustomerSpecList.bxSpecList.Selected(0) = True
Forms.frmCustomerSpecList.Customer = Customer

Set cnn = Nothing
Set rst = Nothing
End Function

Function SaveSpectoTable(strCustomer As String, strCustomerSpec As String, strCustomerPartNumber As String)

On Error GoTo ErrorHandler


    Dim cnn As ADODB.Connection
    Dim rst As New ADODB.Recordset

    
    Set cnn = CurrentProject.Connection
    rst.Open "tblCustomerSpecs", cnn, adOpenKeyset, adLockOptimistic
    
    
    
    With rst
        .AddNew
            !CustomerSpec = strCustomerSpec
            !CustomerPartNumber = strCustomerPartNumber
            !Customer = strCustomer
        .Update
    End With
    
    Set rst = Nothing
    Set cnn = Nothing
    

Error_Exit:
Exit Function
    
ErrorHandler:
'MsgBox (strCustomerPartNumber & " is already associated with " & strCustomerSpec), vbCritical
Resume Error_Exit


    
    


End Function

Function FindBurnListing(ListingCompany As String, Document As String)


Dim strPath As String
Dim Shex As Object
Dim sPath As String
Dim File As String
     
Set Shex = CreateObject("Shell.Application")

If ListingCompany = "ETL" Then
    strPath = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\Product Engineering\UL-ETL\ETL\ETL Listings\Safety\"
  
ElseIf ListingCompany = "UL" Then
    strPath = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\Product Engineering\UL-ETL\UL\UL Listings\"


ElseIf ListingCompany = "MSHA" Then
    strPath = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\Product Engineering\MSHA\MSHA Listings\"

ElseIf ListingCompany = "CPR" Then
        strPath = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\Product Engineering\CPR\CPR Listings\"

End If

    sPath = strPath & "\" & Document & "*"
    File = Dir(sPath)
    
    If IsNull(File) Or File = "" Then
        MsgBox "No file located.", vbCritical
        Exit Function
    End If
    
    strPath = strPath & File
    
    Call OpenCustomerSpec(strPath, File)

End Function

Function UpdateBurnListingFile(NewListingFile As String, NewListingPath, FileType As String, ListingCompany As String)



    Dim strPath As String
    Dim sPath As String
    Dim ArchivePath As String
    Dim File As String
    Dim OldFileName As String
    Dim NewFileName As String
    
    'SpecPath = Forms.frmCreateItem.FileList.ItemData(0)
    Call AddListingReport(NewListingFile, ListingCompany)

    
        
    If ListingCompany = "ETL" Then
        strPath = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\Product Engineering\UL-ETL\ETL\ETL Listings\Safety\"


    ElseIf ListingCompany = "UL" Then
        strPath = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\Product Engineering\UL-ETL\UL\UL Listings\"

    
    ElseIf ListingCompany = "MSHA" Then
        strPath = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\Product Engineering\MSHA\MSHA Listings\"
        
    ElseIf ListingCompany = "CPR" Then
        strPath = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\Product Engineering\UL-ETL\CPR"

    End If

        ArchivePath = strPath & "Archive\"
        sPath = strPath & NewListingFile & "*"
        File = Dir(sPath)
    
    
    'Finds if the path exists
    If Not Dir(strPath) = vbNullString Then
          'Find if the file does not exist
        If Dir(sPath) = vbNullString Then
            'Creates new file if one does not exist
            FileCopy NewListingPath, strPath & NewListingFile & FileType
        Else
            OldFileName = File
            NewFileName = NewListingFile & "_" & Format(Date, "Long Date") & FileType
            'Debug.Print ArchivePath & NewFileName
            If Dir(ArchivePath & NewFileName) = vbNullString Then 'Finds if archived version exists
                Name strPath & OldFileName As ArchivePath & NewFileName 'Creates new file name with date if archived version does not exist
            End If
            FileCopy NewListingPath, strPath & NewListingFile & FileType 'creates New Listing report
        End If
    Else

    End If

    
End Function

Public Sub AddListingReport(Report$, ListingCompany$)

Dim cnn As New ADODB.Connection
Set cnn = New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim provstr As String, qry$


' Specify the OLE DB provider.
cnn.Provider = "sqloledb"


' Specify connection string on Open method.
provstr = cOracleExtractServerConnection
cnn.Open provstr

'TODO Normalize the dbo.tblBurnListingDocument table to have reports/documents unique by combination with the Listing Company
'qry = "  SELECT ListingCompany, ListingDocument FROM dbo.tblBurnListingDocument WHERE ListingCompany =  '" & ListingCompany & "' AND ListingDocument = '" & Report & "'"

qry = "  SELECT ListingCompany, ListingDocument FROM dbo.tblBurnListingDocument WHERE ListingDocument = '" & Report & "'"
Debug.Print qry

rst.Open (qry), cnn, adOpenKeyset, adLockOptimistic

If Not rst.EOF Then
    Exit Sub
End If


    With rst
        .AddNew
            !ListingCompany = ListingCompany
            !ListingDocument = Report
        .Update
    End With
    

rst.Close
cnn.Close


End Sub



Function SaveDocument(DocumentName As String, DocPath As String, FileType As String, NewPath) As Boolean



    Dim NewFile As String
    Dim strResponse As String

    NewFile = DocumentName & FileType
    
    On Error GoTo ErrorHandler


    
    'Finds if the path exists
    If Not Dir(NewPath) = vbNullString Then
          'Find if the file does not does not exist
        If Dir(NewPath & NewFile) = vbNullString Then
            'Creates new file if one does not exist
            FileCopy DocPath, NewPath & NewFile
        Else
            'Asks user if they would like to write over the file
            strResponse = MsgBox("File already exists. Write over file?", vbYesNo + vbCritical, "Override File")
            If strResponse = vbYes Then
                'Writes over existing file
                FileCopy DocPath, NewPath & NewFile
            ElseIf strResponse = vbNo Then
               Exit Function
            End If
        End If
    Else
        'Creates new path if one does not exist
        MkDir (NewPath)
        'Creates new file in the file path
        FileCopy DocPath, NewPath & NewFile
    End If
    
    SaveDocument = True
    
ErrorExit:
    Exit Function

ErrorHandler:
    SaveDocument = False
    Resume ErrorExit

    
            

End Function


Function ShowCableFamilyDocuments(FamilyID As Integer)

    Dim rst As New ADODB.Recordset
    Dim cnn As ADODB.Connection
    Dim strPath As String
    Dim File As Variant
    Dim Path As String
    
    Set cnn = CurrentProject.Connection
    
    
    Path = "\\naa.fujikurausa.com\dfs\SPB\Teams\Cable_Tech\NPRConstructions\"
    
    rst.Open ("Select * FROM tblCableFamilyDocs WHERE FamilyID = " & FamilyID & ""), cnn, adOpenKeyset, adLockOptimistic
    
    Forms.frmCableFamilyViewDocument.bxDocList.rowSource = ""
    
    If rst.RecordCount < 1 Then
        MsgBox "No NPR Records Found", vbCritical
        Exit Function
    End If
    
    'Do Until rst.EOF
    
        strPath = Path & FamilyID & "." & "*"
        
        File = Dir(strPath)
           While (File <> "")
              'If InStr(file, CustomerSpec) > 0 Then
                'strFileName = Path & "\" & File & vbCrLf & strFileName
                'strFile = strPathSpec & File
                'strFile = File
                Forms.frmCableFamilyViewDocument.bxDocList.AddItem File
                 'MsgBox "found " & vbCrLf & strFileName
                 'Exit Function
              'End If
             File = Dir
          Wend
          
        rst.MoveNext
      
    'Loop
      
    Set cnn = Nothing
    Set rst = Nothing
    
    'Do Until rst.EOF
    '    Forms.frmCableFamilyViewDocument.bxDocList.AddItem rst!Document
    '    rst.MoveNext
    'Loop
    '


End Function


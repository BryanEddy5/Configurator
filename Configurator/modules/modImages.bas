Option Compare Database
Option Explicit
Public Const ImageIsometricPath As String = "\\naa.fujikurausa.com\dfs\SPB\Apps\prod\PREMISE\Premise Cable Images\Isometric View\"
Public Const ImageCrossSectionPath As String = "\\naa.fujikurausa.com\dfs\SPB\Apps\prod\PREMISE\Premise Cable Images\Cross-Sectional View\"
Private m_imagesCablePath
Enum ImageTypes
Isometric = 2
CrossSection = 1
End Enum
Private mlImageGroupID As Integer


Public Property Get ImagesCablePath() As String
    If m_imagesCablePath = "" Then m_imagesCablePath = Core_PathwayManager.GetPathway(CableImages)
    ImagesCablePath = m_imagesCablePath
End Property




Public Function getImageCableColor(Image As String) As Variant
Dim hold() As String
Dim i As Variant
Dim ImageName$

ImageName$ = CreateObject("Scripting.FileSystemObject").GetBaseName(Image)

'Iterate through all images in recordset in an attempt to find the color indicator
If ImageName Like "*-*" Then
    hold = Split(ImageName, "-")
    For Each i In hold
        If Len(i) = 1 Then
            getImageCableColor = getColorByColorLetter((i))
            If Not IsNull(getImageCableColor) Then Exit Function
            
        End If
    
    Next
    
End If


    

End Function

Public Function AddImagesToRecordset(Images As FileDialog, ImageType%, IsArmored As Boolean) As Boolean
''add records to the local images table''
Dim rst As New ADODB.Recordset
Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim fileName$, Color$, path

On Error GoTo ErrorHandler

AddImagesToRecordset = True
Set cnn = CurrentProject.connection
rst.Open ("CutSheetImages_Local"), cnn, adOpenDynamic, adLockOptimistic



For Each path In Images.SelectedItems
    fileName = Dir(path)
    Color = getImageCableColor(fileName)
    'fileDir = Left(path, InStrRev(path, "\"))
    
    With rst
        .AddNew
            !ImageName = fileName
            !Color = Color
            !ImageTypeID = ImageType
            !Armored = IsArmored
            !ImagePath = path
        .Update
        
    End With

Next

ErrorExit:
rst.Close
cnn.Close
Exit Function

ErrorHandler:
If Err.Number = -2147217887 Then
    AddImagesToRecordset = False
    rst.CancelUpdate
    Resume Next
Else
    MsgBox Err.Number & " " & Err.Description
    Resume ErrorExit
End If

End Function

Public Function CheckImageGroupReferenceExists(Images As DAO.Recordset, GroupID%) As Boolean
'Check to see if the image is already assigned for that group
Dim rst As New ADODB.Recordset
Dim cnn As New ADODB.connection
Dim qry$, existingItems$
Dim provstr As String



' Specify the OLE DB provider.
cnn.Provider = "SQLOLEDB"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr

'On Error GoTo ErrorHandler

CheckImageGroupReferenceExists = False
Images.MoveFirst
Do Until Images.EOF


    qry = "SELECT * FROM cutsheet.ImageGroupReference  where ImageGroupID = '" & GroupID _
    & "' AND ImageTypeID = " & Images!ImageTypeID & " AND  Armored = " & CInt(Int(Images!Armored)) & " AND Color = '" & Images!Color & "'"
    'Debug.Print qry
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    If rst.RecordCount > 0 Then
        CheckImageGroupReferenceExists = True
        existingItems = existingItems & vbCrLf & Images!ImageName & "Color: " & Images!Color & ", Armored: " & Images!Armored & " , Image type: " & Images!ImageTypeID
        Images.Delete
        
    End If
    
    rst.Close
    Images.MoveNext
Loop

If existingItems <> "" Then
    MsgBox "The following cable images already exist for this group" & vbCrLf & existingItems
End If

End Function
Public Function UpdateBaseImageID()
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryUpdateImageGroupID"

End Function


Public Function AddImages(Images As DAO.Recordset, GroupID%) As Boolean

''Add images to the cutsheet images table''
''***Notify user that the image for that group already exists. Ask if they would like to overwrite it?****''
Dim rst As New ADODB.Recordset
Dim cnn As New ADODB.connection
Dim fileExt$, qry$
Dim provstr As String

mlImageGroupID = GroupID




' Specify the OLE DB provider.
cnn.Provider = "SQLOLEDB"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr

On Error GoTo ErrorHandler

Images.MoveFirst
Do Until Images.EOF
fileExt = Split(Images!ImageName, ".")(UBound(Split(Images!ImageName, ".")))
    

    qry = "SELECT * FROM cutsheet.Images where ImageName = '" & Images!ImageName & "' AND ImageTypeID = " & Images!ImageTypeID
    'Debug.Print qry
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
        With rst
            If .RecordCount > 0 Then
                If MsgBox("An image with this name already exists, do you want to add it to this group?", vbYesNo) = vbYes Then
                    GoTo addImage
                End If
            Else
                .AddNew
                    !ImageName = Images!ImageName
                    !ImageTypeID = Images!ImageTypeID
                    !Color = Images!Color
                    !Armored = Images!Armored
                    !ImageFileType = "." & fileExt
                .Update
addImage:
                If Not SaveImageFile(Images, rst) Then
                    rst.Delete
                Else
                    Call AddImageToGroupReference(rst, GroupID%)
                    Images.Delete
                End If
            End If
                    
            
        End With
    rst.Close
    Images.MoveNext
Loop


ErrorExit:

Exit Function

ErrorHandler:
If Err.Number = -2147217887 Then
    rst.CancelUpdate
    Resume Next
Else
    MsgBox Err.Number & " " & Err.Description
    Resume ErrorExit
End If

End Function


Private Function SaveImageFile(Images As DAO.Recordset, TableSavedImages As ADODB.Recordset) As Boolean
''Copy image files to the appropriate location''
Dim fileDestination$
Dim Destination, errorMessage$

On Error GoTo ErrorHandler

SaveImageFile = False

''****New logic for saving as Image ID*****'''
Destination = ImagesCablePath & TableSavedImages!ImageID & TableSavedImages!ImageFileType

If CommonFunctions.FileExists(Destination) Then
    If MsgBox("File " & Images!ImageName & " already exists.  Would you like to overwrite it?", vbYesNo) = vbNo Then
        SaveImageFile = True
        Exit Function
    End If
End If

FileCopy Images!ImagePath, Destination
SaveImageFile = True

ErrorExit:
Exit Function

ErrorHandler:
If Err.Number = 70 Then
    errorMessage = "Error on saving image " & Images!ImageName & ". Permission to save to folder has been denied. "
Else

    errorMessage = "Error on saving image " & Images!ImageName & ".  Please ensure your network has not been interrupted. "
End If
MsgBox errorMessage
SaveImageFile = False
Resume ErrorExit


End Function


Private Function AddImageToGroupReference(Record As ADODB.Recordset, ImageGroupID%) As Boolean
''Associate the loaded images with the group id''
Dim rst As New ADODB.Recordset
Dim qry$
Dim provstr As String

AddImageToGroupReference = False


Set rst = Core_DataAccess.GetAdodbRecordset(qry)
    With rst
        .AddNew
            !ImageGroupID = ImageGroupID%
            !ImageID = Record!ImageID
            !ImageTypeID = Record!ImageTypeID
            !Color = Record!Color
            !Armored = Record!Armored
        .Update
    End With
    
rst.Close
AddImageToGroupReference = True


ErrorExit:
Exit Function

ErrorHandler:
If Err.Number = -2147217887 Then
    AddImageToGroupReference = False
    rst.CancelUpdate
    Resume Next
Else
    MsgBox Err.Number & " " & Err.Description
    Resume ErrorExit
End If

End Function

Public Function AddImageGroupID(GroupName, groupDesc) As Integer
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim qry$
'Dim DateCreated As Date
Dim provstr As String


On Error GoTo ErrorHandler:
' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
qry = "SELECT * FROM CutSheet.ImageGroup"
rst.Open qry, cnn, adOpenKeyset, adLockOptimistic


    With rst
        .AddNew
            !ImageGroupName = GroupName
            !ImageGroupDesc = groupDesc
        .Update
        AddImageGroupID = !ImageGroupID
    End With



ErrorExit:
rst.CancelUpdate
rst.Close
cnn.Close
Exit Function

ErrorHandler:
If Err.Number = -2147217873 Then
    MsgBox ("This image group already exists.  Transaction has been cancelled"), vbCritical
Else
    MsgBox Err.Number & " " & Err.Description
End If

Resume ErrorExit




End Function
Function setImageGroupID(ImageGroupID%)

    mlImageGroupID = ImageGroupID


End Function

Function GetCableImageID() As String

    GetCableImageID = mlImageGroupID

End Function




Public Function UpdateImageGroupDescription(ImageGroupDesc$, ImageGroupID%) As Boolean
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim qry$
'Dim DateCreated As Date
Dim provstr As String


On Error GoTo ErrorHandler:
' Specify the OLE DB provider.
cnn.Provider = "sqloledb"
UpdateImageGroupDescription = False
' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
qry = "SELECT * FROM CutSheet.ImageGroup where ImageGroupID = " & ImageGroupID%
rst.Open qry, cnn, adOpenKeyset, adLockOptimistic


With rst
    If Not .BOF And Not .EOF Then
        
        If .Supports(adUpdate) Then
            !ImageGroupDesc = ImageGroupDesc
            .Update
            UpdateImageGroupDescription = True
            
        End If
    End If

End With

ErrorExit:
rst.Close
cnn.Close
Exit Function

ErrorHandler:
    MsgBox Err.Number & " " & Err.Description

Resume ErrorExit





End Function




Sub FilesAndFolders()
    Dim FSO As Object 'FileSystemObject
    Dim fldStart As Object 'Folder
    Dim fld As Object 'Folder
    Dim Mask As String
    
    On Error GoTo Error_Resume

    Set FSO = CreateObject("scripting.FileSystemObject") ' late binding
    'Set fso = New FileSystemObject 'or use early binding (also replace Object types)

    Set fldStart = FSO.GetFolder("\\naa.fujikurausa.com\dfs\SPB\Apps\prod\PREMISE\Premise Cable Images") ' <-- use your FileDialog code here

    Mask = "*.jpg"
    Debug.Print fldStart.path & "\"
    ListFiles fldStart, Mask
    For Each fld In fldStart.SubFolders
        ListFiles fld, Mask
        ListFolders fld, Mask
    Next
    
Error_Resume:
    Resume Next
End Sub


Sub ListFolders(fldStart As Object, Mask As String)
    Dim fld As Object 'Folder
    For Each fld In fldStart.SubFolders
        If fld.name Like "*cross*" Then
            'Debug.Print fld.Path & "\"
            ListFiles fld, Mask
            ListFolders fld, Mask
        End If
    Next

End Sub



Sub MatchImagesInTable()
' Initialize variables.
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim strSQL As String
'Dim DateCreated As Date
Dim provstr As String

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr


strSQL = "SELECT * FROM cutsheet.CutSheet_Images"
rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
End Sub

Function GetAllImages()



Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset
Dim ImageType(2) As Integer
Dim item As Variant
Dim Color As Variant
Dim i As Integer
Color = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "L", "K")


ImageType(0) = 1
ImageType(1) = 2

On Error GoTo ErrorHandler


' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
sQry = "SELECT * FROM vFindImages_AllBases"

rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
    Do Until rst.EOF
        For i = 0 To 1
            For Each item In Color
                TestCableImage rst!DesignCode, rst!CablePasses, i, Nz(rst!CableFamily), rst!Base, item
            Next
        Next
    rst.MoveNext
    Loop
    
rst.Close
cnn.Close

Exit Function


ErrorHandler:
Resume Next



End Function

Function ImageToTable(ImageName As String, ImageType As String, Location As String, Base As String)
Dim rst As New ADODB.Recordset
Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection

Set cnn = CurrentProject.connection

rst.Open ("tblMissingImages"), cnn, adOpenDynamic, adLockOptimistic

With rst
    .AddNew
        !ImageName = ImageName
        !ImageType = ImageType
        !Location = Location
        !Base = Base
    .Update
End With



End Function

Function FoundImageToTable(ImageName As String, ImageType As Integer, Location As String, Base As String, Color As String)
Dim rst As New ADODB.Recordset
Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection

Set cnn = CurrentProject.connection

rst.Open ("tblFoundImages"), cnn, adOpenDynamic, adLockOptimistic

With rst
    .AddNew
        !ImageName = ImageName
        !ImageType = ImageType
        !Location = Location
        !Base = Base
        !Color = Color
    .Update
End With



End Function

Function TestCableImage(sDesignCode As String, iLocation As Integer, ByVal ImageType As Integer, CableFamily As String, Base As String, ByVal Color As String) As String

Dim ImageIsometricPath As String
Dim ImageCrossSectionPath As String


    Dim sFolder As String
    Dim sFilePath As String
    Dim sImage As String
    Dim File As String
    Dim sNumFibersPerTube As String
    Dim sNumSubs As String
    Dim sFibersPerBundle As String
    Dim sCableConfiguration As String
    Dim sCableFamily As String
    Dim sPath As String
    Dim sBaseImage As String
    Dim ImageStyle As String
    
    Dim CableProperties As New clsCableConstruction
    
    CableProperties.Load (Base)

    If ImageType = 2 Then
        sPath = ImageIsometricPath
        ImageStyle = "Isometric View"
    ElseIf ImageType = 1 Then
        sPath = ImageCrossSectionPath
        ImageStyle = "Cross-Sectional View"

    End If



    sCableFamily = CableFamily

    Select Case sCableFamily
        Case "K"
            sCableFamily = "C"
    End Select


    'Set ItemData = RetrieveItemData(sItem)
        'sDesignCode = ItemData.Item("DesignCode")
    'sFolder = FindItemData(sItem, "DesignTypeGroupID") & "\" & FindItemData(sItem, "DesignTypeID")
    sFilePath = sPath & sFolder & "\"

   ' If iLocation = 2 Then
   If CableProperties.NumPositions = 0 Then
    CableProperties.NumPositions = 1
    End If
            sNumFibersPerTube = HexLetter(CableProperties.FiberCount / (CableProperties.NumPositions - CableProperties.NumFillers))
            sNumSubs = HexLetter(CableProperties.NumPositions)
            
            If CableProperties.FiberCount < 12 Then
                sFibersPerBundle = HexLetter(CableProperties.FiberCount)
            ElseIf CableProperties.FiberCount Mod 12 = 0 Then
                sFibersPerBundle = HexLetter(12)
            ElseIf CableProperties.FiberCount Mod 8 = 0 Then
                sFibersPerBundle = HexLetter(8)
            Else
                sFibersPerBundle = Hex(CableProperties.FiberCount Mod 12)
            End If
            
            sCableConfiguration = "-" & sNumFibersPerTube & sNumSubs & sFibersPerBundle
'
'    Else
'
'            sCableConfiguration = ""
'
'    End If
    sBaseImage = sDesignCode & "-" & Color
    sImage = sBaseImage & ".jpg"
    File = Dir(sFilePath & sImage)
    If File = vbNullString Then
            'sImage = sBaseImage
            sImage = sBaseImage & sCableConfiguration & ".jpg"
        File = Dir(sFilePath & sImage)
            If File = vbNullString Then
                sImage = sBaseImage & "-" & CableProperties.FiberCount & ".jpg"
                File = Dir(sFilePath & sImage)
                If File = vbNullString Then
                'MsgBox "Missing Cable Image. Error: 2220", vbCritical
                'Debug.Print sImage & " " & ImageStyle & " Not Found"
                'Call ImageToTable(sImage, ImageStyle, "Not Found", Base)
                    Exit Function
                End If
            End If

    End If

    'Debug.Print File
    'Debug.Print sImage & " " & ImageStyle & " Found"
     Call FoundImageToTable(sImage, ImageType, sFilePath, Base, Color)


End Function


Function GetNewImageNameFromFileName(fileName As String, Folder As String) As String


Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset
Dim ImageType%



On Error GoTo ErrorHandler

If Folder Like "*Isometric*" Then
    ImageType = 2
Else
    ImageType = 1
End If




' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr
sQry = "SELECT * FROM CutSheet.Images WHERE ImageName = '" & fileName & "' AND ImageTypeID = " & ImageType

rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic

    If rst.EOF Then
        Debug.Print "Image is missing from table: " & fileName & " FileType" & ImageType
        Exit Function
    End If
    
        GetNewImageNameFromFileName = rst!ImageID & rst!ImageFileType
        
    
 
    
rst.Close
cnn.Close

Exit Function


ErrorHandler:
Resume Next



End Function
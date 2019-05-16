Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Public Base As String, ItemNo As String, Customer As String, PID As String, Reason As String, Fiber1 As String, Fiber2 As String, Fiber3 As String, CableColor As String, OracleTBDesc As String
Public Location As Integer, TBLetterIndicator As Integer
Public ActiveConstruction As Boolean, ReleasedDesign As Boolean, CableItemActive As Boolean, TightBuffered As Boolean, RibbonHighCure As Boolean, ComponentSelection As Boolean
Private mlPrinted As Boolean
Private mlPrintLine1 As String, mlPrintLine2 As String, mlPrintLine3 As String, mlPrintLine4 As String, mlPrintItemNo As Boolean, mlPrintLotNo As Boolean
Private mlItemNumber As String, mlListingLabel As String, mlSpecSheetApprovalStatus As String

Public Property Get ItemNumber() As String

ItemNumber = mlItemNumber

End Property

Public Property Let ItemNumber(ItemNumber$)

mlItemNumber = ItemNumber

End Property




Public Function Load(ItemNumber As String) As Boolean

On Error GoTo HandleError:

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr

sQry = "SELECT * FROM [vCableQualities] WHERE ItemNumber = '" & ItemNumber & "'"

rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
    Load = False
    mlItemNumber = ItemNumber

    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If
    
            ' Note use of Nz function in case of Null values in the table
            Me.Base = Nz(!Base)
            Me.ItemNo = Nz(!ItemNo)
            Me.PID = Nz(!PID)
            Me.Customer = Nz(!Customer)
            Me.Reason = Nz(!Reason)
            Me.Fiber1 = Nz(!Fiber1)
            Me.Fiber2 = Nz(!Fiber2)
            Me.Fiber3 = Nz(!Fiber3)
            Me.CableColor = Nz(!CableColor)
            Me.Location = Nz(!CablePasses)
            Me.ActiveConstruction = Nz(!Active)
            Me.ReleasedDesign = Nz(!ReleasedDesign)
            Me.ReleasedDesign = Nz(!ReleasedDesign)
            Me.CableItemActive = Nz(!CableItemActive)
            Me.TightBuffered = Nz(!TightBuffered)
            Me.RibbonHighCure = Nz(!RibbonHighCure)
            Me.OracleTBDesc = Nz(!OracleTBDesc)
            Me.TBLetterIndicator = Nz(!TBLetterIndicator)
            Me.ComponentSelection = Nz(!ComponentSelection)
            mlPrintLine1 = Nz(!PrintLine1)
            mlPrintLine2 = Nz(!PrintLine2)
            mlPrintLine3 = Nz(!PrintLine3)
            mlPrintLine4 = Nz(!PrintLine4)
            mlPrintItemNo = Nz(![Print Item No])
            mlPrintLotNo = Nz(![Print Reel No])
        .Close
    End With
    Load = True
    
    Call loadSpecSheetApprovalStatus
    
'Call SubUnit

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function

Public Property Let IsPrinted(IsPrinted As Boolean)

mlPrinted = IsPrinted

End Property

Public Property Get IsPrinted() As Boolean

IsPrinted = mlPrinted

End Property

Public Property Get getPrintString() As String
Dim PrintLine2 As String, PrintLine3 As String


If mlPrintItemNo Then
    PrintLine2 = mlPrintLine1 & " " & mlPrintLine2 & " " & Me.ItemNo
Else
    PrintLine2 = mlPrintLine1 & " " & mlPrintLine2
End If

If mlPrintLotNo Then
    PrintLine3 = mlPrintLine3 & " Lot No. ###### " & mlPrintLine4
Else
    PrintLine3 = mlPrintLine3 & " " & mlPrintLine4
End If


getPrintString = PrintLine2 & " " & PrintLine3

End Property



Private Sub loadSpecSheetApprovalStatus()


On Error GoTo HandleError:

Dim cnn As New ADODB.connection
Set cnn = New ADODB.connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset
Dim criteria As String

criteria = "(CableUnion.Customer LIKE 'AFL STANDARD%' OR CableUnion.Customer LIKE 'AFL GENERIC%' OR  CableUnion.Oracle LIKE 'DN%')"

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = Core_SqlConnections.ConnectionString
cnn.Open provstr

sQry = "SELECT Status FROM dbo.[vfindcutsheet] WHERE Oracle = '" & mlItemNumber & "'"

rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic



    With rst
        If .EOF Then
            mlSpecSheetApprovalStatus = "No Spec Available"
            GoTo Done
        End If

            ' Note use of Nz function in case of Null values in the table
            mlSpecSheetApprovalStatus = Nz(!Status, "No Spec Available")
        .Close
    End With


'Call SubUnit

Done:
    Exit Sub
HandleError:
    MsgBox "Error " & Err.Number & " while retrieving approval status:" & vbCrLf & Err.Description, vbCritical
    Resume Done

End Sub

Public Property Get getSpecSheetApprovalStatus() As String
getSpecSheetApprovalStatus = mlSpecSheetApprovalStatus

End Property

Public Function setListingLabel(ListingCompany As String, UsListing As Boolean, CanadaListing As Boolean)

If mlPrinted = False Or ListingCompany = "NONE" Then
    Exit Function
End If

If UsListing And CanadaListing Then
    mlListingLabel = "(" & ListingCompany & ")c(" & ListingCompany & ")"
ElseIf UsListing Then
    mlListingLabel = "(" & ListingCompany & ")"
ElseIf CanadaListing Then
    mlListingLabel = "c(" & ListingCompany & ")"
Else
    mlListingLabel = ""
End If

End Function


Public Property Get getListingLabel() As String

    getListingLabel = mlListingLabel

End Property
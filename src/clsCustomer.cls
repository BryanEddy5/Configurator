VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Option Explicit

Public Customer  As String
Private ExistingCustomer As Boolean
Private mlCustomerSpecialNotes As String
Private m_CustomerProhibitedNotes, m_CustomerID


Public Property Get CustomerProhibitedNotes() As String

CustomerProhibitedNotes = m_CustomerProhibitedNotes

End Property

Public Property Get CustomerID() As String

CustomerID = m_CustomerID

End Property


Public Function Load(Customer As String) As Boolean


On Error GoTo HandleError:

Dim cnn As New ADODB.Connection
Set cnn = New ADODB.Connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = cPremiseServerConnection
cnn.Open provstr
sQry = "SELECT * FROM dbo.vCustomers WHERE Customer = '" & Customer & "'"

rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
    Load = False


    With rst
        If .RecordCount = 0 Then
            GoTo Done
        End If
    
            ' Note use of Nz function in case of Null values in the table
            Me.Customer = Nz(!Customer)
            mlCustomerSpecialNotes = Nz(!CustomerSpecialNotes)
            m_CustomerID = Nz(!CustomerID)
            
        .Close
    End With
    Load = True
    ExistingCustomer = Load

Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
'
End Function



Function IsProhibitedDesign(Base As String) As Boolean  ' Read only

On Error GoTo HandleError:

Dim cnn As New ADODB.Connection
Set cnn = New ADODB.Connection
Dim provstr As String
Dim sQry As String
Dim rst As New ADODB.Recordset

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"

' Specify connection string on Open method.
provstr = cPremiseServerConnection
cnn.Open provstr
sQry = "SELECT * FROM CustomerProhibitedConstructions WHERE Customer = '" & Me.Customer & "' AND Base = '" & Base & "'"

rst.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
 
    IsProhibitedDesign = False

    If rst.EOF Then
        Exit Function
    End If
        m_CustomerProhibitedNotes = rst!Comments
'        With rst
'            If .RecordCount = 0 Then
'                GoTo Done
'            End If
'            .Close
'        End With
    IsProhibitedDesign = True
rst.Close
cnn.Close
Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading item information:" & vbCrLf & Err.Description, vbCritical
    Resume Done
End Function

Property Get IsExistingCustomer() As Boolean      ' Read only
    IsExistingCustomer = ExistingCustomer
End Property


'---------------------------------------------
Public Function AddNewCustomer(Customer As String) As Boolean
'---------------------------------------------
' Purpose:  Save the object to the database table (New document for first time)
' Requires: All properties to be valid
' Returns:  True if successful; false if fails
    
    On Error GoTo HandleError
    
    AddNewCustomer = False
    
Dim cnn As New ADODB.Connection
Set cnn = New ADODB.Connection
Dim provstr As String
Dim sQry As String
Dim rs As New ADODB.Recordset

' Specify the OLE DB provider.
cnn.Provider = "sqloledb"
Me.Customer = Customer

' Specify connection string on Open method.
provstr = cPremiseServerConnection
cnn.Open provstr
sQry = "SELECT * FROM Customers WHERE Customer = '" & Customer & "'"

rs.Open (sQry), cnn, adOpenKeyset, adLockOptimistic
    
    With rs
        .AddNew
            SaveFieldsToDb rs
        .Update
        
        '.Bookmark = .LastModified
        'mlOrderID = !OrderID   ' Get ID from new record
        .Close
    End With
    Set rs = Nothing
    
    AddNewCustomer = True
    
Done:
    Exit Function
    
HandleError:
    MsgBox "Error " & Err.Number & " in Create" & vbCr & Err.Description, vbCritical
    Resume Done
    
End Function

'-------------------------------------------
Private Sub SaveFieldsToDb(rs As ADODB.Recordset)
'-------------------------------------------
' Purpose:  Save properties that can be updated to database
' Requires: All properties to be valid;
'           Open recordset <rs> for writing to
' Returns:  --
    With rs
        !Customer = Me.Customer
    End With

End Sub

Property Get getCustomerNotes() As String

    getCustomerNotes = mlCustomerSpecialNotes

End Property





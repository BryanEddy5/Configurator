Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'======================================
' Class: clsOrder1
'======================================
' A more advanced Order object with Load, Create, Update and Delete methods

'************************COPYRIGHT NOTICE*************************
' Copyright (C) 2000 DI Management Services Pty Ltd,
' Sydney Australia www.di-mgt.com.au. All rights reserved.
' This code was originally written by David Ireland.
' You are free to use it in any application
' provided this copyright notice is left intact.
' If you use it, or found it useful, or can suggest an improvement
' please let us know at code@di-mgt.com.au.
'*****************************************************************

Option Compare Database
Option Explicit

Private Const scTABLE As String = "tblOrders"

Private mlOrderID As Long
Public CustomerID As Long
Public item As String
Public Qty As Long
Public UnitPrice As Currency

'---------------------------------------------
Property Get OrderID() As Long      ' Read only
    OrderID = mlOrderID
End Property

'-------------------------------------------
Public Function Load(ID As Long) As Boolean
'-------------------------------------------
' Purpose:  Load the object with ObjectID=<ID> from the database
' Requires: Record to exist
' Returns:  True if successful; false if fails
'
    On Error GoTo HandleError
    
    Dim rs As DAO.Recordset
    Dim sQry As String
    
    Load = False
    
    sQry = "SELECT * FROM " & scTABLE & " WHERE ([OrderID]=" & ID & ");"
    
    Set rs = CurrentDb().OpenRecordset(sQry, dbOpenForwardOnly)
    With rs
        If .RecordCount = 0 Then
            MsgBox "Cannot find Order with ID = " & ID, vbCritical
            GoTo Done
        End If
        
        ' Note use of Nz function in case of Null values in the table
        mlOrderID = Nz(!OrderID)
        Me.CustomerID = Nz(!CustomerID)
        Me.item = Nz(!item)
        Me.Qty = Nz(!Qty)
        Me.UnitPrice = Nz(!UnitPrice)
        
        .Close
    End With
    Load = True
    
Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " while loading Order:" & vbCrLf & Err.Description, vbCritical
    Resume Done

End Function

'---------------------------------------------
Public Function Create() As Boolean
'---------------------------------------------
' Purpose:  Save the object to the database table (New document for first time)
' Requires: All properties to be valid
' Returns:  True if successful; false if fails
    
    On Error GoTo HandleError
    
    Create = False
    
    Dim rs As DAO.Recordset
    
    Set rs = CurrentDb().OpenRecordset(scTABLE, dbOpenDynaset)
    With rs
        .AddNew
        SaveFieldsToDb rs
        .Update
        
        .Bookmark = .LastModified
        mlOrderID = !OrderID   ' Get ID from new record
        .Close
    End With
    Set rs = Nothing
    
    Create = True
    
Done:
    Exit Function
    
HandleError:
    MsgBox "Error " & Err.Number & " in Create" & vbCr & Err.Description, vbCritical
    Resume Done
    
End Function

'-------------------------------------------
Private Sub SaveFieldsToDb(rs As DAO.Recordset)
'-------------------------------------------
' Purpose:  Save properties that can be updated to database
' Requires: All properties to be valid;
'           Open recordset <rs> for writing to
' Returns:  --
    With rs
        !CustomerID = Me.CustomerID
        !item = Me.item
        !Qty = Me.Qty
        !UnitPrice = Me.UnitPrice
    End With

End Sub

'---------------------------------------------
Public Function Update() As Boolean
'---------------------------------------------
' Purpose:  Save the revised Document to the database table
' Requires: Record must already exist in database
' Returns:  True if successful; false if fails
'
    On Error GoTo HandleError
    
    Update = False
    
    Dim rs As DAO.Recordset
    Dim sQry As String
    
    sQry = "SELECT * FROM " & scTABLE & " WHERE ([OrderID]=" & Me.OrderID & ");"
    Set rs = CurrentDb().OpenRecordset(sQry, dbOpenDynaset)
    With rs
        If .RecordCount = 0 Then
            .Close
            Err.Raise 40000, , "Cannot find OrderID=" & Me.OrderID
            GoTo Done
        End If
        
        .Edit
        SaveFieldsToDb rs
        .Update
        
        .Close
    End With
    Set rs = Nothing
    
    Update = True
    
Done:
    Exit Function
    
HandleError:
    MsgBox "Error " & Err.Number & " in Update" & vbCr & Err.Description, vbCritical
    Resume Done
    
End Function

'---------------------------------------------
Public Function Delete() As Boolean
'---------------------------------------------
' Purpose:  Delete this object from the db
' Requires: Calling program to have checked that deletion is OK
' Returns:  True if successful; false if fails
'
    On Error GoTo HandleError

    Dim sQry As String
    
    Delete = False
    
    If Me.OrderID <= 0 Then
        GoTo Done
    End If
    
    sQry = "DELETE * FROM " & scTABLE & " WHERE ([OrderID]=" & Me.OrderID & ");"
    CurrentDb().Execute sQry, dbFailOnError
    
    Delete = True
Done:
    Exit Function
HandleError:
    MsgBox "Error " & Err.Number & " in Delete" & vbCr & Err.Description, vbCritical
    Resume Done

End Function
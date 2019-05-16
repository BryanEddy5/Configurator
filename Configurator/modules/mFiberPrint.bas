Option Compare Database
Option Explicit
Private mlDesignCode As String, FiberFamily As String, mlTbLetter As String
Private mlFiber As New Collection
Private m_VendorSpecific As Boolean
Private bFiberExtension As Boolean
Private mlFiberCount As Integer, mlTBTypeID As Integer
Private m_fiberVendor$


Public Property Get FiberVendor()
    FiberVendor = m_fiberVendor
End Property


Public Property Get isVendorSpecificPrint() As Boolean

isVendorSpecificPrint = m_VendorSpecific

End Property




Private Sub UnitTestFiberPrint()
Dim Fiber1 As New clsFiber, bCorningPrint As Boolean, DesignCode As String, TBTypeID As Integer, Fiber2 As New clsFiber, Fiber3 As New clsFiber, _
cFiber  As New Collection


Fiber1.Load ("fbr00389")
Fiber1.setFiberCount = 12
Fiber1.setFiberLetter = "X"
cFiber.Add Fiber1

Fiber2.Load ("fbr00414")
'Fiber2.setFiberCount = 24
'Fiber2.setFiberLetter = "C"
cFiber.Add Fiber2

Fiber3.Load ("FBR00387")
'Fiber3.setFiberCount = 12
'Fiber3.setFiberLetter = "C"
cFiber.Add Fiber3

bCorningPrint = 0
DesignCode = "GQ048C2029CB-8SM9/16SM4/12MMN"
FiberExtension (DesignCode)
TBTypeID = 4





Debug.Print getFiberPrint(cFiber, DesignCode, TBTypeID) & vbCrLf
'Call getMultipleFiberCount(cFiber, DesignCode)
'
'Debug.Print Fiber1.getFiberCount
'
'Debug.Print Fiber2.getFiberCount
'
'Debug.Print Fiber3.getFiberCount

'loop through and print all fiber counts
'For each



End Sub


Public Function getFiberPrint(ByRef Fiber As Collection, DesignCode As String, TBTypeID As Integer) As String

On Error GoTo ErrorCatcher:

Set mlFiber = Fiber
m_VendorSpecific = False
mlDesignCode = DesignCode
mlTBTypeID = TBTypeID
mlFiberCount = Int(Mid(DesignCode, 3, 3))
FiberFamily = Left(DesignCode, 1) & Mid(DesignCode, 3, 3)
mlTbLetter = Mid(DesignCode, 12, 1)

'If Mid(mlDesignCode, 9, 1) = "U" Then
'    Exit Function
'End If


getMultipleFiberCount


getFiberPrint = StandardPrint()

If Len(getFiberPrint) > 70 Then
    MsgBox "Fiber Print is greater than 70 characters." & vbCrLf & "Get commercial approval for print.", vbCritical
End If
'
ErrorExit:
Exit Function

ErrorCatcher:
MsgBox Err.Number & " " & Err.Description
Resume ErrorExit



End Function

Private Function getFiberCount() As String

Select Case mlTbLetter
    Case "R", "S"
    getFiberCount = mlFiberCount & " FIBER SPIDERWEB RIBBON (R)"
    Case Else
        Select Case FiberFamily
            Case "Z002"
                getFiberCount = "ZIPCORD"
            Case "S001"
                getFiberCount = "SIMPLEX"
            Case Else
                getFiberCount = mlFiberCount & " FIBER"
        End Select
End Select


End Function


Private Function StandardPrint() As String

    Dim lngKt As Integer
    Dim i As Integer
    Dim sPrint As String

    lngKt = 0
    i = 0
    
    'If design is a furcation then use the print string and exit
    If mlTBTypeID = 34 Or mlTBTypeID = 5 Then
        StandardPrint = "FURCATION TUBE"
        Exit Function
    ElseIf mlTBTypeID = 32 Then
        Exit Function
    End If
    
    'iterate through the fiber collection
    lngKt = mlFiber.Count
    For i = 1 To lngKt
        sPrint = sPrint & FiberPrint(mlFiber(i), i, lngKt)
    Next
    

    
    'Final print string with the fiber count concatenated
   StandardPrint = sPrint & " " & getFiberCount()

    'empties collection
    lngKt = mlFiber.Count
    For i = 1 To lngKt
        mlFiber.Remove 1
    Next
    


End Function

Private Function FiberPrint(ByRef Fiber As clsFiber, Position As Integer, Count As Integer) As String
    
    Dim VendorCheck As Boolean
    
    VendorCheck = False
        
        'If Vendor is corning specific then check for corning indiactors
        'If Fiber.Vendor = "corning" Then
            FiberPrint = getCorningPrint(Fiber, Count)
            VendorCheck = m_VendorSpecific
        'Else
            'VendorCheck = False 'If the function determines it is not corning specific then flag it so it will run a generic print
        'End If
        
        If VendorCheck = False Then
            FiberPrint = GenericPrint(Fiber, Count) ' Creates generic fiber print
        End If
    
    
        'If the
        If Count <> Position Then
            FiberPrint = FiberPrint & " / "     'Adds seperator for each fiber type if there is more than one fiber
        End If


End Function

Private Function GenericPrint(ByRef Fiber As clsFiber, Count As Integer) As String
        
        'Formats the print string dependent on the amount of fibers
        If Count > 1 Then
            GenericPrint = Fiber.getFiberCount & " X " & Fiber.Micron
        Else
            GenericPrint = Fiber.Micron
        End If
        
        'Generic Print string
        Select Case Fiber.Mode
            Case "SM"
                If Fiber.BIF = "Yes" And Fiber.getFiberLetter = "9" Then           'Generic SM BIF
                    GenericPrint = GenericPrint & " BIF"
                ElseIf Fiber.BIF = "Yes" Then
                    GenericPrint = GenericPrint & " BIF " & Fiber.ITU    'Non-Generic SM BIF to include ITU standard
                Else
                    GenericPrint = GenericPrint
                End If
            Case Else
                If Fiber.BIF = "Yes" Then
                    GenericPrint = GenericPrint & " " & Fiber.OM & " BIF"      'BIF MM fiber
                Else
                    GenericPrint = GenericPrint & " " & Fiber.OM               'Non-BIF MM Fiber
                End If
        End Select
        
End Function

Public Function getCorningPrint(ByRef Fiber As clsFiber, Count As Integer) As String
    Dim cnn As New ADODB.connection
    Set cnn = New ADODB.connection
    Dim provstr As String
    Dim rst As New ADODB.Recordset

    Dim strSQL As String
    Dim strItemExt() As String
    Dim i As Integer
    Dim FiberLetter, FiberExt As String
    

    ' Specify the OLE DB provider.
    cnn.Provider = "sqloledb"
    
    ' Specify connection string on Open method.
    provstr = Core_SqlConnections.ConnectionString
    cnn.Open provstr

    m_VendorSpecific = False
    FiberLetter = Fiber.getFiberLetter
    
    strItemExt() = Split(mlDesignCode, "-")
    
    
        'Select the proper formatting if there is only 1 fiber or multiple
        If Count > 1 Then
            getCorningPrint = Fiber.getFiberCount & " X "
        Else
            getCorningPrint = getCorningPrint
        End If
    
    
            'Check if any of the extensions are corning specific
            Do While i <= UBound(strItemExt) Or m_VendorSpecific = True
                FiberExt = strItemExt(i)
            
        
                strSQL = "SELECT * FROM vVendorSpecificPrint WHERE FiberExtensions = '" & FiberExt & "' AND FiberLetter = '" & FiberLetter & "'"
                rst.Open (strSQL), cnn, adOpenKeyset, adLockOptimistic
                If rst.EOF = True Then
                    m_VendorSpecific = False
                Else
                    m_VendorSpecific = True
                    Exit Do
                End If
                rst.Close
                i = i + 1
            Loop
            

        'if none of of the extensions are corning specific then check for corning print with fiber extensions of "NONE"
        If m_VendorSpecific = False Then
        
            FiberExt = "NONE"
            strSQL = "SELECT * FROM vVendorSpecificPrint WHERE FiberLetter = '" & FiberLetter & "' and FiberExtensions = '" & FiberExt & "'"
                        rst.Open (strSQL), cnn, adOpenKeyset, adLockOptimistic
            If rst.EOF = True Then
                m_VendorSpecific = False
                Exit Function
            Else
                    m_VendorSpecific = True
            End If
            
        End If
        
        If m_VendorSpecific = False Then
            Exit Function
        End If
        
        

    
    
    
    getCorningPrint = getCorningPrint & rst!VendorSpecificPrint
    
    m_fiberVendor = UCase(rst!VendorName)

'    If mlDesignCode Like "*CRG+*" Then
'        getCorningPrint = Replace(getCorningPrint, "OM4", "OM4+")
'    End If

    rst.Close
    cnn.Close

End Function



'Identifies if the Item has a fiber Extension indicating a certain fiber to use
Public Function FiberExtension(DesignCode) As Integer

Dim cnn As ADODB.connection
Dim rst As New ADODB.Recordset
Dim intExtFound As Boolean
Set cnn = CurrentProject.connection
Dim strSQL As String
Dim sExtension As String
Dim i As Variant
Dim Ext$()

'sExtension = Mid(DesignCode, 14, 30)
Ext = Split(DesignCode, "-")


FiberExtension = 0

    If sExtension Like "*bif*" And Not sExtension Like "*bif+*" Then
        FiberExtension = 2
        intExtFound = True
        Exit Function
    End If


For Each i In Ext

        strSQL = "SELECT DISTINCT FiberExtension "
        strSQL = strSQL & " FROM tbl_Fibers INNER JOIN tblFiberExtensions ON tbl_Fibers.Fiber = tblFiberExtensions.Fiber "
        strSQL = strSQL & "WHERE  active <> 0 AND FiberExtension = '" & i & "'"
        rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
        
        If Not rst.EOF Then
        
            FiberExtension = 1
                    
            sFiberExtension = rst!FiberExtension
            
            bFiberExtension = True
            
            intExtFound = bFiberExtension
        End If
        rst.Close
Next

        


End Function

Public Sub getMultipleFiberCount()
Dim Fiber As Collection
Dim DesignCode As String
Dim ExtensionHold() As String
Dim FiberCountHold() As String
Dim i As Variant
Dim k As Integer
k = 1

Set Fiber = mlFiber
DesignCode = mlDesignCode

If Mid(DesignCode, 3, 3) Like "*@*" Then
    Exit Sub
End If

'If there is only 1 fiber then return the fiber letter and fiber count
If Fiber.Count = 1 Then
    Fiber.item(1).setFiberLetter = Mid(DesignCode, 6, 1)
    Fiber.item(1).setFiberCount = Int(Mid(DesignCode, 3, 3))
    Exit Sub
Else

'Segment the design code up into all extensions
ExtensionHold = Split(DesignCode, "-")

'Iterate through each extension to determine the extensions that is for split fibers
For Each i In ExtensionHold

    If i Like "*/*" Then
        FiberCountHold = Split(i, "/")
    End If

Next



If UBound(FiberCountHold) > 0 Then

    For Each i In FiberCountHold
        Fiber(k).setFiberLetter = Right(i, 1)       'Get fiber letter for each multi fiber segment
        Fiber(k).setFiberCount = MultiFiberCount(i) 'Get the fiber count for each fiber
        k = k + 1
    Next

End If

    
End If



End Sub

Function MultiFiberCount(s As Variant) As Variant
    ' Variables needed (remember to use "option explicit").   '
    Dim RetVal As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    RetVal = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If IsNumeric(Mid(s, i, 1)) = False And i > 1 Then
            Exit Function
        ElseIf IsNumeric(Mid(s, i, 1)) Then
            MultiFiberCount = MultiFiberCount + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    MultiFiberCount = RetVal
End Function
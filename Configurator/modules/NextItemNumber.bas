Option Explicit
Option Compare Database

Function NextItemNo(strItemNo As String, bSub As Boolean) As String


    Dim cnn As ADODB.connection
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim strCableLetter As String
    Dim strLastOracleNumber As String
    Dim strNewOracleNumber As String
    Dim strTightBuffer As String
    Dim strPremisPrefix$
    Set cnn = CurrentProject.connection
    
    strCableLetter = Left(strItemNo, 1)
    strTightBuffer = Mid(strItemNo, 12, 1)
    
    'For Tight buffered ribbon, this determines that a "PW00XXX" item will be built
    If Left(strItemNo, 2) = "TK" And Not Left(strItemNo, 5) = "tn001" Then
        strCableLetter = "W"
    End If
    
    
    
    
    
    If bSub = False Then
    
        If Left(strItemNo, 2) = "RB" Then
            'If MsgBox("Do you want to create an RBN item? ", vbYesNo) = vbYes Then
                strPremisPrefix = "RBN"
'            Else
'            strPremisPrefix = "P" & strCableLetter
'            End If
        Else
        
            strPremisPrefix = "P" & strCableLetter
            
        End If
        
        If strCableLetter = "R" And strTightBuffer = "8" Then
        
            strSQL = "SELECT TOP 1 * FROM qryItemNoJoin "
            strSQL = strSQL & "WHERE [Oracle] like '" & strPremisPrefix & "%'"
            strSQL = strSQL & "ORDER BY Oracle DESC "
            
        Else
        
            strSQL = "SELECT TOP 1 * FROM qryItemNoJoin "
            strSQL = strSQL & "WHERE [Oracle] like '" & strPremisPrefix & "%' AND [Oracle] not like 'RBN01%'"
            strSQL = strSQL & "ORDER BY Oracle DESC "
        End If
    
    Else
    
        strPremisPrefix = "P" & strCableLetter
    
        strSQL = "SELECT TOP 1 * FROM qryItemNoJoin_Subs "
        strSQL = strSQL & "WHERE [Oracle] like '" & strPremisPrefix & "%' AND [Oracle] not like 'RBN01%'"
        strSQL = strSQL & "ORDER BY Oracle DESC "
    End If
    
    
    
    rst.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    
    'rst.MoveFirst
    
    
    'Debug.Print rst!Oracle
    
    strLastOracleNumber = rst!Oracle
    
    If strPremisPrefix = "RBN" Then
        strNewOracleNumber = Left(rst!Oracle, 3) & Format(Int(Mid(rst!Oracle, 4, 5) + 1), "00000")
        'Debug.Print strNewOracleNumber
    Else
        strNewOracleNumber = Left(rst!Oracle, 2) & Format(Int(Mid(rst!Oracle, 3, 5) + 8), "00000")
        'Debug.Print strNewOracleNumber
    End If
    
    NextItemNo = strNewOracleNumber
    
    
    rst.Close
    Set cnn = Nothing

Forms!frmCreateItem!UnitSeries = Left(strNewOracleNumber, 7)

End Function
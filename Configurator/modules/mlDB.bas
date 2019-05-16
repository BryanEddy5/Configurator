Option Explicit
Option Compare Database

Function ChangeProperty(strPropName As String, varPropType As Variant, varPropValue As Variant) As Integer
    Dim dbs As Object, prp As Variant
    Const conPropNotFoundError = 3270
 
    Set dbs = CurrentDb
    On Error GoTo Change_Err
    dbs.Properties(strPropName) = varPropValue
    ChangeProperty = True
 
Change_Bye:
    Exit Function
 
Change_Err:
    If Err = conPropNotFoundError Then    ' Property not found.
        Set prp = dbs.CreateProperty(strPropName, _
            varPropType, varPropValue)
        dbs.Properties.Append prp
        Resume Next
    Else
        ' Unknown error.
        ChangeProperty = False
        Resume Change_Bye
    End If
End Function
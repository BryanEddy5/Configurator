Option Compare Database
Option Explicit

Public Enum pathway
    CableImages = 1
    CustomerSpecs
    ETL
    UL
    MSHA
    CPR
    NPR
End Enum


Public Function GetPathway(pathway As pathway) As String

    Dim sql As String
    Dim rst As New ADODB.Recordset
    Dim path As String, separator As String
    
    sql = "SELECT Pathway FROM Core.Pathways WHERE PathId = " & pathway
    
    Set rst = Core_DataAccess.GetAdodbRecordset(sql)
    
    If rst.EOF Then
        MsgBox "An error occurred attempting to get pathway"
        Exit Function
    
    End If
    
    path = rst!pathway
    separator = "\"
    
    If Right(path, 1) <> separator Then path = path & separator
    
#If DEV Then
    path = path & "Dev" & separator
#End If
    
    GetPathway = path
    


End Function
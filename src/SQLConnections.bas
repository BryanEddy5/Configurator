Attribute VB_Name = "SQLConnections"
Option Explicit
Option Compare Database

'Global Const cOracleExtractServerConnection = "Server=NAASPB-PRD04\SQL2014;Database=OracleExtracts;Trusted_Connection=yes"
Global Const cOracleExtractServerConnection = "Server=NAASPB-PRD04\SQL2014;Database=Premise;Trusted_Connection=yes"

Global Const cPremiseServerConnection = "Server=NAASPB-PRD04\SQL2014;Database=Premise;Trusted_Connection=yes"

'Global Const cPremiseServerConnection = "Provider=SQLNCLI11.1;Integrated Security=SSPI;Persist Security Info=False;User ID="";Initial Catalog=Premise;Data Source=NAASPB-PRD04\SQL2014;Initial File Name="";Server SPN="""

Function SQLUpdateConnection()
Dim tdf As TableDef
Dim db As Database

On Error GoTo ErrorHandler

    Set db = CurrentDb

    For Each tdf In CurrentDb.TableDefs
        If tdf.Connect Like "*OracleExtracts*" Or tdf.Connect Like "*Premise*" Then
  
            
            'Debug.Print Replace(tdf.Connect, "OracleExtracts", "Premise")
            
            'tdf.Connect = Replace(tdf.Connect, "OracleExtracts", "Premise")
            
            tdf.RefreshLink

           Debug.Print tdf.Connect
           'Debug.Print tdf.Connect
        End If
    Next
    
ErrorHandler:
'Debug.Print tdf.Name & tdf.SourceTableName & tdf.Connect
Resume Next



End Function




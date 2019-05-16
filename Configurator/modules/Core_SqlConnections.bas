Option Explicit
Option Compare Database


Private Const m_dbDev = "Database=Premise_DEV;"
Private Const m_dbProd = "Database=Premise;"
Private Const m_server = "Server=NAASPB-PRD04\SQL2014;"
Private Const m_connectionType = "Trusted_Connection=yes"



Private Property Get PremiseServerConnection() As String

    Dim server, db$, cnn$
    
    server = m_server
    
#If DEV Then
    db = m_dbDev
#Else
    db = m_dbProd
#End If
    
    cnn = m_connectionType
    
    PremiseServerConnection = server & db & cnn

End Property

Public Property Get ConnectionString() As String

ConnectionString = PremiseServerConnection

End Property



Function SQLUpdateConnection()
Dim tdf As TableDef
Dim db As Database
Dim ProdDb$, DevDb$, NewConnection$, OldConnection$

On Error GoTo ErrorHandler

    Set db = CurrentDb
    ProdDb = m_dbProd
    DevDb = m_dbDev
    
    NewConnection = ProdDb
    OldConnection = DevDb

#If DEV Then
    NewConnection = DevDb
    OldConnection = ProdDb
#End If

    For Each tdf In CurrentDb.TableDefs
    
        If tdf.connect Like "*" & OldConnection & "*" Then
  
            
            
            'Debug.Print Replace(tdf.Connect, "OracleExtracts", "Premise")
            
            tdf.connect = Replace(tdf.connect, OldConnection, NewConnection)
            
            tdf.RefreshLink

           Debug.Print tdf.connect & " " & tdf.name
           'Debug.Print tdf.Connect
        End If
    Next
    
ErrorHandler:
'Debug.Print tdf.Name & tdf.SourceTableName & tdf.Connect
Resume Next



End Function
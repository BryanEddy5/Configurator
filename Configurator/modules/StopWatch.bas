Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



'***********  Start Code  ****************
Private mlngStart As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub StartTimer()
    mlngStart = GetTickCount
End Sub

Public Function EndTimer() As String
    Dim EndTime As Long
    Dim ms As Integer
    Dim s As Long

    EndTime = (GetTickCount - mlngStart)
    ms = EndTime Mod 1000
    s = (EndTime - ms) / 1000
    If s >= 1 Then
        EndTimer = s & "s, " & ms & "ms"
    Else
        EndTimer = ms & "ms"
    End If

End Function
' ***********  End Code  ****************

'
'
'' ***********  Start Code  **************
'Dim sw As StopWatch
'Set sw = New StopWatch
'sw.StartTimer
'
'' Do whatever you want to time here
'
'Debug.Print "That took: " & sw.EndTimer & "milliseconds"
'' *********  End Code  *************
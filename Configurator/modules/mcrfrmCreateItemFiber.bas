Option Compare Database
Option Explicit
'------------------------------------------------------------
' mcrfrmCreateItemFiber
'
'------------------------------------------------------------
Function frmCreateItemFiber()
On Error GoTo mcrfrmCreateItemFiber_Err

    With CodeContextObject
        .OM = Null
        .BIF = Null

        DoCmd.Requery "BIF"
        DoCmd.Requery "OM"

        .PrintLine2 = Null
        DoCmd.Requery "PrintLine2"
    End With


mcrfrmCreateItemFiber_Exit:
    Exit Function

mcrfrmCreateItemFiber_Err:
    MsgBox Error$
    Resume mcrfrmCreateItemFiber_Exit

End Function
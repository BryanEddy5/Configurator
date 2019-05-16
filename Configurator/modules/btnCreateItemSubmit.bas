Option Compare Database
Option Explicit
'------------------------------------------------------------
' mcrbtnCreateItemSubmit
'
'------------------------------------------------------------
Function mcrbtnCreateItemSubmit()
On Error GoTo mcrbtnCreateItemSubmit_Err


    
    DoCmd.OpenForm "frmWait", acNormal, "", "", , acNormal
    ' Clear Basic Cable Hold Table
    DoCmd.OpenQuery "qryNewBasicCableDelete", acViewNormal, acEdit
    ' Clear Basic Construction Hold Table
    DoCmd.OpenQuery "qryNewPrductConstructionHoldClear", acViewNormal, acEdit
    ' Clear Unit Series Hold Table
    DoCmd.OpenQuery "qrytblUnitSeriesCopyClear", acViewNormal, acEdit
    
    
    If (Forms!frmCreateItem!Location = "1") Then
    End If
    If (Forms!frmCreateItem!Location = "2") Then
        ' creates new item in Basic (Cable/Sheath), unit series, and subunits
        Call mcrCreateBasicCableSheath
    End If
            

mcrbtnCreateItemSubmit_Exit:
    Exit Function

mcrbtnCreateItemSubmit_Err:
    MsgBox Error$
    Resume mcrbtnCreateItemSubmit_Exit

End Function
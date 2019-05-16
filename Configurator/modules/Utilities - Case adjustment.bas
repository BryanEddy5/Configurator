Option Compare Database
Option Explicit

' this next is intended to be keyed into the Menu system
' so that automatic 'case' adjustment can be disabled/enabled

Dim glb_AutoCase_inActive As Boolean
'

Public Function Proper(objVal As Variant) As String
'*******************************************
'Purpose:   Impose 'standardised' case conventions on the given string
'           the initial letter of each word is capitalised ...
'           except that
'           1.  Word is defined as a sequence of characters bounded by such as " ,.-'()/"
'           2.  Common linking words (on, and, the etc) are not capitalised
'           3.  Words of the form  Mac... and Mc... are also capitalised after the Mac or Mc
'           4.  Some testing is done to avoid obvious mistakes on place names
'                   ... 'shire', 'Betws-y-coed', 'Hampden-in-Arden' ...
'
'Author:    CD
'Date:      Oct 8, 1994, 11:46
'Called by: General in-line code - most freq from ProperForm
'*******************************************

    Dim i%, lenf%, new_word As Boolean, force_up As Boolean, wk$, wka$, strVal$

    Proper = ""
    If Nz(objVal, "") = "" Then Exit Function
    strVal = LCase(Trim(objVal))
    lenf = Len(strVal)
    ReDim A$(lenf)
    For i = 1 To lenf
        A$(i) = Mid(strVal, i, 1)
        Next i
    i = 0: wk$ = ""
    wka$ = strVal & "......"      ' Avoid the silly VB fault in re testing literals against shorter strings
    new_word = True
    While i < lenf
        i = i + 1
        Select Case A$(i)
            Case " ": new_word = True
            Case ".": new_word = True
            Case ",": new_word = True
            Case "-": new_word = True
            Case "'": new_word = True
            Case "(": new_word = True
            Case ")": new_word = True
            Case "/": new_word = True
            Case Else
                    If new_word Then
                        A$(i) = UCase(A$(i))        ' force initial character of new word to UPPER case
                        '   check for the celtic surnames ... we don't really need to go overboard on all the
                        '   words in English starting with mac - this routine is usually used in the context
                        '   of name and address, company name, job title etc., so hit as few as possible
                        force_up = True
                        If Mid(strVal, i, 3) = "mac" Then
                            Select Case Proper_Word$(wka$, i)
                                Case "machin":          force_up = False
                                Case "machine":         force_up = False
                                Case "machines":        force_up = False
                                Case "mace":            force_up = False
                                Case "maces":           force_up = False
                                Case "macey":           force_up = False
                                Case "mackerel":        force_up = False
                                Case "macclesfield":    force_up = False
                                End Select
                            If force_up Then
                                If lenf >= i + 3 Then A$(i + 3) = UCase(A$(i + 3))
                                End If
                            Else
                            If Mid(strVal, i, 2) = "mc" Then
                                If lenf >= i + 2 Then A$(i + 2) = UCase(A$(i + 2))
                                Else
                                If Proper_UndoCapital(wka$, i) Then A$(i) = LCase(A$(i))
                                End If
                            End If
                        new_word = False
                        End If
            End Select
        wk$ = wk$ & A$(i)
        Wend

    Proper = wk$
    End Function

Private Function Proper_UndoCapital(wka$, i%)
'*******************************************
'Purpose:   The character in "wka$" pointed to by "i" is the 1st
'               character of a new word in the string.
'           It has so far been forced to upper case.
'           Can we find a reason to turn it back to lower case.
'
'Author:    CD
'Date:      Oct 8, 1994, 11:46
'Called by: Proper
'*******************************************

    Proper_UndoCapital = True
    If i > 1 Then
        ' So we are not at beginning of given string
        ' ... do we really need the 1st letter of word to be capitalised

        '  1st - are we interested in the previous character ?
        If i > 2 Then
            Select Case Mid(wka$, i - 2, 4)
                Case "c/o ": Exit Function          ' trap out on the 'o' part
                End Select
            End If
        Select Case Mid(wka$, i - 1, 3)
            Case " a ": Exit Function
            Case "'s ": Exit Function
            Case "'s.": Exit Function
            Case "'s,": Exit Function
            Case "'s;": Exit Function
            Case "-y-": Exit Function
            End Select
        If Mid(wka$, i - 1, 4) = "-in-" Then Exit Function
        If Mid(wka$, i - 1, 6) = "-shire" Then Exit Function

        '   OK - then what about the word itself
        '   ... but not as 1st word of phrase, 'cos we're in the middle of string
        Select Case Mid(wka$, i, 3)
            Case "as ": Exit Function
            Case "of ": Exit Function
            Case "on ": Exit Function
            Case "in ": Exit Function
            Case "le ": Exit Function
            End Select
        Select Case Mid(wka$, i, 4)
            Case "and ": Exit Function
            Case "for ": Exit Function
            Case "the ": Exit Function
            End Select
        If Mid(wka$, i, 5) = "upon " Then Exit Function
        End If
    
    '   and what if the word _is_ at the beginning of the phrase
    If Mid(wka$, i, 4) = "etc " Then Exit Function
    If Mid(wka$, i, 4) = "etc." Then Exit Function
    If Mid(wka$, i, 4) = "c/o " Then Exit Function      ' trap out on the 'c' part
    If Mid(wka$, i, 3) = "on-" Then Exit Function
    If Mid(wka$, i, 3) = "plc" Then Exit Function
    If Mid(wka$, i, 3) = "de " Then Exit Function
    If Mid(wka$, i, 3) = "en " Then Exit Function
    If Mid(wka$, i, 3) = "en-" Then Exit Function
    If Mid(wka$, i, 2) = "d'" Then Exit Function
    Proper_UndoCapital = False
    End Function

Private Function Proper_Word$(txt$, st%)
'*******************************************
'Purpose:   Return the 'word' in txt$ starting with the st%'th character
'           'word' here is simplistic ... only contains letters
'
'Author:    CD
'Date:      Oct 8, 1994, 11:46
'Called by: Proper
'*******************************************

    Dim i%
    i = st
    Do
        If i > Len(txt$) Then Exit Do
        Select Case Mid$(txt$, i, 1)
            Case "a" To "z"
            Case Else: Exit Do
            End Select
        i = i + 1
        Loop
    Proper_Word$ = Mid$(txt$, st, i - st)
    End Function

Public Function ProperForm()
'*******************************************
'Purpose:   Capitalise current value of control on form ... unless
'           the only change is a change of case, in which event
'           leave value alone, else user will get PO'd
'
'Author:    CD
'Date:      Oct 8, 1994, 11:46
'Called by: AfterUpdate event of 'text-box'-type control
'           =ProperForm()
'*******************************************

    ' ... This is a function only for Access syntax reasons,
    '     ...  so that it may be called from Property sheet
    '   it returns no value
    
    Dim ctrl As control, doIt As Boolean
    
    On Local Error Resume Next
    
    Set ctrl = Screen.ActiveControl     '   ... saves bothering to pass arg
    doIt = Not IsNull(ctrl) And Not glb_AutoCase_inActive
    If Err <> 0 Then doIt = False: Err.Clear
    If doIt Then
        If Not IsNull(ctrl.OldValue) Then
            If ctrl.OldValue = ctrl Then
                ' the only thing that can have changed is the "case" of the letters, so leave it alone
                doIt = False
                End If
            End If
        If doIt Then ctrl = Proper(ctrl)
        End If
    Set ctrl = Nothing
    End Function
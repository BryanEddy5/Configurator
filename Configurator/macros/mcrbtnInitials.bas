Version =196611
ColumnsShown =2
Begin
    Condition ="[Forms]![frmSwitchBoardInitials]![Initials] Is Null"
    Action ="MsgBox"
    Argument ="Please enter your initials before continuing."
    Argument ="-1"
    Argument ="1"
    Argument ="Enter Initials"
End
Begin
    Condition ="[Forms]![frmSwitchBoardInitials]![Initials] Is Null"
    Action ="StopMacro"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![frmSwitchBoard]![Initials]"
    Argument ="[Forms]![frmSwitchBoardInitials]![Initials]"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmSwitchBoardInitials"
    Argument ="0"
End

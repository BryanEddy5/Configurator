Version =196611
ColumnsShown =2
Begin
    Condition ="False"
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![NewOracle#]"
    Argument ="Left([Forms]![frmFiberMatchExisting]![Oracle],8)"
End
Begin
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![Exists]"
    Argument ="-1"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![Fiber]"
    Argument ="[Forms]![frmFiberMatchExisting]![Fiber]"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![Jacket]"
    Argument ="[Forms]![frmFiberMatchExisting]![Jacket]"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![Color]"
    Argument ="[Forms]![frmFiberMatchExisting]![Color]"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![Base]"
    Argument ="[Forms]![frmFiberMatchExisting]![Base]"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![ColorChip]"
    Argument ="[Forms]![frmFiberMatchExisting]![ColorChip]"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmFiberMatchExisting"
    Argument ="0"
End
Begin
End
Begin
End
Begin
End
Begin
    Condition ="False"
    Action ="Close"
    Argument ="2"
    Argument ="frmFiberMatchExisting"
    Argument ="0"
End

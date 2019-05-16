Version =196611
ColumnsShown =2
Begin
    Action ="OpenForm"
    Argument ="frmUnitSeriesExisting"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Condition ="[Forms]![frmUnitSeries]![NewUnit] Is Null Or [Forms]![frmUnitSeriesExisting]![Ex"
        "isting]=-1"
    Action ="MsgBox"
    Argument ="Please enter a new Unit Series Number."
    Argument ="-1"
    Argument ="1"
    Argument ="Unit Series"
End
Begin
End
Begin
    Condition ="[Forms]![frmUnitSeriesExisting]![Existing]=0 Or [Forms]![frmUnitSeries]![NewUnit"
        "] Is Not Null"
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![UnitSeries]"
    Argument ="[Forms]![frmUnitSeries]![NewUnit]"
End
Begin
    Condition ="[Forms]![frmUnitSeriesExisting]![Existing]=0 Or [Forms]![frmUnitSeries]![NewUnit"
        "] Is Not Null"
    Action ="Close"
    Argument ="2"
    Argument ="frmUnitSeries"
    Argument ="0"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmUnitSeriesExisting"
    Argument ="2"
End

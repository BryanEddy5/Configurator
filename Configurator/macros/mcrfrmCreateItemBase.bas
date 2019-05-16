Version =196611
ColumnsShown =2
Begin
    Action ="SetValue"
    Argument ="[ListingCompany]"
    Argument ="Null"
End
Begin
    Action ="Requery"
    Argument ="ListingCompany"
End
Begin
    Condition ="[Forms]![frmCreateItem]![Jacket] Is Not Null"
    Action ="SetValue"
    Argument ="[Jacket]"
    Argument ="Null"
End
Begin
    Condition ="[Forms]![frmCreateItem]![Jacket] Is Not Null"
    Action ="Requery"
    Argument ="Jacket"
End
Begin
    Action ="SetValue"
    Argument ="[Color]"
    Argument ="Null"
End
Begin
    Action ="Requery"
    Argument ="Color"
End
Begin
    Action ="SetValue"
    Argument ="[ColorChip]"
    Argument ="Null"
End
Begin
    Action ="Requery"
    Argument ="ColorChip"
End

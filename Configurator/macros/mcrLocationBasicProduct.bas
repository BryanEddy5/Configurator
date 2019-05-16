Version =196611
ColumnsShown =2
Begin
    MacroName ="Forms!frmLocation!BasicCable = -1"
    Condition ="[Forms]![frmMainMenu]![BasicProduct]=-1"
    Action ="SetValue"
    Argument ="[BasicCable]"
    Argument ="0"
End
Begin
    MacroName ="Forms!frmLocation!BasicCable = 0"
    Condition ="[Forms]![frmMainMenu]![BasicProduct]=0"
    Action ="SetValue"
    Argument ="[BasicCable]"
    Argument ="-1"
End
Begin
    MacroName ="Forms!frmLocation!BasicCable = 0"
End

Version =196611
ColumnsShown =0
Begin
End
Begin
    Action ="SetValue"
    Comment ="Turn on footer section"
    Argument ="[Forms]![frmNewPrintExtension].[Section](2).[Visible]"
    Argument ="Yes"
End
Begin
    Action ="GoToControl"
    Argument ="NewPrintExtension"
End
Begin
    Action ="SetValue"
    Comment ="Turn off Yes button"
    Argument ="[Forms]![frmNewPrintExtension]![cmdNo].[Visible]"
    Argument ="No"
End
Begin
    Action ="SetValue"
    Comment ="Turn off No button"
    Argument ="[Forms]![frmNewPrintExtension]![cmdYes].[Visible]"
    Argument ="No"
End

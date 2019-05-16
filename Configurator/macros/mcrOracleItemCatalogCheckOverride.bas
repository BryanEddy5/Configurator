Version =196611
ColumnsShown =2
Begin
End
Begin
    Condition ="[Forms]![frmItemInterface]![frmItemInterfaceCatalog].[Form]![AllowOverride]=0"
    Action ="MsgBox"
    Argument ="Information in this field can not be changed.  If this information is incorrect,"
        " select cancel to end the Auto-Load process."
    Argument ="-1"
    Argument ="1"
    Argument ="EDIT NOT ALLOWED"
End
Begin
    Condition ="[Forms]![frmItemInterface]![frmItemInterfaceCatalog].[Form]![AllowOverride]=0"
    Action ="GoToControl"
    Argument ="[ElementName]"
End
Begin
    Condition ="[Forms]![frmItemInterface]![frmItemInterfaceCatalog].[Form]![AllowOverride]=0"
    Action ="StopAllMacros"
End

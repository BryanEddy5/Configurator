Version =196611
ColumnsShown =2
Begin
    Action ="MsgBox"
    Argument ="Please create new Unit Series Number before continuing."
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="Close"
    Comment ="closes fiber spec"
    Argument ="2"
    Argument ="frmFiberSpec"
    Argument ="0"
End
Begin
    Action ="Close"
    Comment ="closes item copied from"
    Argument ="1"
    Argument ="qryCopyItemMatch"
    Argument ="0"
End
Begin
    Condition ="False"
    Action ="Close"
    Argument ="2"
    Argument ="frmOracleMatch"
    Argument ="0"
End

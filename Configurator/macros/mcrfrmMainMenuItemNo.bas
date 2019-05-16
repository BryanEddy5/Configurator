Version =196611
ColumnsShown =2
Begin
End
Begin
    Action ="OpenForm"
    Argument ="frmOracleMatch"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
End
Begin
End
Begin
    Action ="OpenForm"
    Argument ="frmCreateItem"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Condition ="[Forms]![frmMainMenu]![BasicCable]=-1"
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![Location]"
    Argument ="2"
End
Begin
    Condition ="[Forms]![frmMainMenu]![BasicProduct]=-1"
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![Location]"
    Argument ="1"
End
Begin
    Condition ="[Forms]![frmMainMenu]![Bypass]=0"
    Action ="SetValue"
    Comment ="Set Value in frmCreateItem to show if the item exists or not"
    Argument ="[Forms]![frmCreateItem]![Exists]"
    Argument ="[Forms]![frmOracleMatch]![Exists]"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![ItemNo]"
    Argument ="[Forms]![frmMainMenu]![ItemNo]"
End
Begin
    Condition ="[Forms]![frmOracleMatch]![Exists]=-1 And [Forms]![frmMainMenu]![Bypass]=0"
    Action ="RunMacro"
    Argument ="mcrLoadExistingData"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmOracleMatch"
    Argument ="0"
End
Begin
    Condition ="[Forms]![frmMainMenu]![Bypass]=-1"
    Action ="SetValue"
    Argument ="[Forms]![frmCreateItem]![Exists]"
    Argument ="0"
End

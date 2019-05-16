Version =196611
ColumnsShown =2
Begin
End
Begin
    Action ="OpenQuery"
    Comment ="Clear temp BOM table"
    Argument ="qryItemInterfaceTempBOMClear"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Comment ="Load temp BOM table"
    Argument ="qryItemInterfaceTempBOMLoad"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenForm"
    Argument ="frmBOMEdit"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
End
Begin
    Action ="RunCode"
    Argument ="TBUpdate ()"
End

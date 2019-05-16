Version =196611
ColumnsShown =2
Begin
    Action ="OpenForm"
    Comment ="Open form showing Item to copy and next subunit oracle number"
    Argument ="frmSubUnitNextItemNo"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="OpenForm"
    Argument ="frmFiberSpec"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="OpenQuery"
    Comment ="clear table to hold subunit information"
    Argument ="qryNewBasicConstructionClear"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Argument ="qryNewPrductConstructionHoldClear"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Comment ="Add item to copy's information to table"
    Argument ="qrySubunitAppend"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenForm"
    Argument ="frmTBMaterial"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="OpenForm"
    Argument ="frmSubUnitJacketColor"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="OpenQuery"
    Comment ="Update the information in the table to create new item"
    Argument ="qrySubunitUpdate"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Comment ="Update the TB Material for the subunit [Forms]![frmCreateItem]![TBType]<>\"C\" A"
        "nd [Forms]![frmCreateItem]![TBType]<>\"S\" And [Forms]![frmCreateItem]![TBType]<"
        ">\"B\""
    Argument ="qryfrmCreateItemTBMaterialUpdate"
    Argument ="0"
    Argument ="1"
End
Begin
End
Begin
    Action ="OpenQuery"
    Comment ="Updates the tblUnitSeriesCopy with the newly created item"
    Argument ="qrySubunitNewUnitSeriesUpdate"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Comment ="Appends to Basic Product Hold table"
    Argument ="qryNewBasicConstructionAppend"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="False"
    Action ="OpenQuery"
    Comment ="Append to Basic Product Consctruction"
    Argument ="qrySubUnitAppendBasicConstruction"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmTBMaterial"
    Argument ="0"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmFiberSpec"
    Argument ="0"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmSubUnitJacketColor"
    Argument ="0"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmSubUnitNextItemNo"
    Argument ="0"
End
Begin
    Action ="OpenForm"
    Argument ="frmFinishNewSubunit"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End

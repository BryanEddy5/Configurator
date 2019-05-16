Version =196611
ColumnsShown =2
Begin
    Action ="OpenForm"
    Argument ="frmSpecialPrintMatch"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="OpenForm"
    Comment ="IIf([Forms]![frmCreatItem]![Color] Like \"*black*\",\"SPECIAL: WHITE PRINT\",\"S"
        "tandard black print\")"
    Argument ="frmFiberSpec"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="OpenQuery"
    Comment ="Clears table"
    Argument ="qryNewBasicConstructionClear"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Comment ="Appends Item to Copy from to temp table"
    Argument ="qryNewBasicConstruction"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="False"
    Action ="Close"
    Argument ="1"
    Argument ="qryCopyItemMatch"
    Argument ="0"
End
Begin
    Action ="OpenQuery"
    Comment ="Updates Item to new information in temp table"
    Argument ="qryUpdateNewProductConstruciton"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="[Forms]![frmCreateItem]![ItemNo] Like \"*000@*\""
    Action ="OpenQuery"
    Comment ="Updates information for furcation tube"
    Argument ="qryfrmCreateItemFurcationUpdate"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="RunCode"
    Argument ="TBMaterial_Update () "
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmFiberSpec"
    Argument ="0"
End
Begin
    Condition ="[Forms]![frmCreateItem]![PrintType]<>0"
    Action ="OpenQuery"
    Comment ="Uploads Item to Hold Table"
    Argument ="qryNewBasicConstructionAppend"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="[Forms]![frmCreateItem]![Unprinted]=-1"
    Action ="OpenQuery"
    Comment ="Updates Item to unprinted information in temp table"
    Argument ="qryNewProductBasicUnprintedUpdate"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="[Forms]![frmCreateItem]![Unprinted]=-1"
    Action ="OpenQuery"
    Comment ="Upload unprinted new item to Hold Table"
    Argument ="qryNewProductBasicUnprintedAppend"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Comment ="Upload Items in Hold Table to Basic Construction"
    Argument ="qryLOADBasicProductConstruction"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="False"
    Action ="OpenQuery"
    Argument ="qryOpenNewProduct"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="False"
    Action ="OpenQuery"
    Argument ="qryNewBasicProductFindFamily"
    Argument ="0"
    Argument ="1"
End
Begin
End
Begin
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmSpecialPrintMatch"
    Argument ="0"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmWait"
    Argument ="0"
End
Begin
    Action ="OpenForm"
    Argument ="frmNewSubUnit"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End

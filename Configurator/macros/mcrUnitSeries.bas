Version =196611
ColumnsShown =2
Begin
    Action ="OpenForm"
    Argument ="frmFiberMode"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="OpenQuery"
    Comment ="Clear Table"
    Argument ="qrytblUnitSeriesCopyClear"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Comment ="Append to temporary Unit Series"
    Argument ="qryCableSheathUnitSeriesCopyAppend"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="OpenQuery"
    Comment ="Update Unit Series to Chosen"
    Argument ="qryUpdateUnitSeries"
    Argument ="0"
    Argument ="1"
End
Begin
End
Begin
End
Begin
    Condition ="[Forms]![frmCreateItem]![SubX]=1"
    Action ="RunCode"
    Comment ="Runs code to create sequentially colored subunits"
    Argument ="frmSubunitX () "
End
Begin
    Condition ="[Forms]![frmCreateItem]![SubX]=1"
    Action ="StopMacro"
    Comment ="Stops Macro"
End
Begin
End
Begin
    Action ="OpenForm"
    Comment ="Find if the Unit Series Subunit Exists"
    Argument ="frmSubUnitExists"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Condition ="[Forms]![frmSubUnitExists]![Exists]=-1"
    Action ="OpenQuery"
    Comment ="Clears tlbUnitSeriesCopy"
    Argument ="qryUnitSeriesSubunitClear"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="[Forms]![frmSubUnitExists]![Exists]=-1"
    Action ="OpenQuery"
    Comment ="Appends the Unit Series to copy from"
    Argument ="qryUnitSeriesSubunitAppend"
    Argument ="0"
    Argument ="1"
End
Begin
    Condition ="[Forms]![frmSubUnitExists]![Exists]=-1"
    Action ="OpenQuery"
    Comment ="Updates the Unit Series copy with new subunit"
    Argument ="qryUnitSeriesCopyItemNoUpdate"
    Argument ="0"
    Argument ="1"
End
Begin
End
Begin
    Condition ="False"
    Action ="MsgBox"
    Argument ="Create New Subunit"
    Argument ="-1"
    Argument ="3"
End
Begin
    Condition ="[Forms]![frmSubUnitExists]![Exists]=0"
    Action ="RunMacro"
    Comment ="Run Macro to create new subunit"
    Argument ="mcrCreateNewSubUnit"
End
Begin
    Condition ="False"
    Action ="OpenQuery"
    Comment ="Append tblUnitSeriesCopy to Unit Series table"
    Argument ="qryUnitSeriesAppend"
    Argument ="0"
    Argument ="1"
End
Begin
    Action ="Close"
    Comment ="Closes form for finding if subunit exists"
    Argument ="2"
    Argument ="frmSubUnitExists"
    Argument ="0"
End
Begin
    Action ="Close"
    Argument ="2"
    Argument ="frmFiberMode"
    Argument ="0"
End

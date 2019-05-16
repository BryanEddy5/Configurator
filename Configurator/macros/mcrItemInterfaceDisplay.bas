Version =196611
ColumnsShown =2
Begin
End
Begin
    Condition ="[Forms]![frmItemInterface]![OrganizationCode]=\"000\""
    Action ="SetValue"
    Comment ="frmItemInterfaceCatalog - Yes"
    Argument ="[Forms]![frmItemInterface]![frmItemInterfaceCatalog].[Form].[Visible]"
    Argument ="Yes"
End
Begin
    Condition ="[Forms]![frmItemInterface]![OrganizationCode]=\"000\""
    Action ="SetValue"
    Comment ="frmItemInterfaceCategory - Yes"
    Argument ="[Forms]![frmItemInterface]![frmItemInterfaceCategory].[Form].[Visible]"
    Argument ="Yes"
End
Begin
    Condition ="[Forms]![frmItemInterface]![OrganizationCode]=\"000\""
    Action ="SetValue"
    Comment ="frmItemInterfaceBOM - No"
    Argument ="[Forms]![frmItemInterface]![frmItemInterfaceBOM].[Form].[Visible]"
    Argument ="No"
End
Begin
    Condition ="[Forms]![frmItemInterface]![OrganizationCode]=\"000\""
    Action ="SetValue"
    Comment ="frmItemInterfaceRouting - No"
    Argument ="[Forms]![frmItemInterface]![frmItemInterfaceRouting].[Form].[Visible]"
    Argument ="No"
End
Begin
End
Begin
    Condition ="[Forms]![frmItemInterface]![OrganizationCode]<>\"000\""
    Action ="SetValue"
    Comment ="frmItemInterfaceCatalog - No"
    Argument ="[Forms]![frmItemInterface]![frmItemInterfaceCatalog].[Form].[Visible]"
    Argument ="No"
End
Begin
    Condition ="[Forms]![frmItemInterface]![OrganizationCode]<>\"000\""
    Action ="SetValue"
    Comment ="frmItemInterfaceCategory - yes"
    Argument ="[Forms]![frmItemInterface]![frmItemInterfaceCategory].[Form].[Visible]"
    Argument ="No"
End
Begin
    Condition ="[Forms]![frmItemInterface]![OrganizationCode]<>\"000\""
    Action ="SetValue"
    Comment ="frmItemInterfaceBOM - Yes"
    Argument ="[Forms]![frmItemInterface]![frmItemInterfaceBOM].[Form].[Visible]"
    Argument ="Yes"
End
Begin
    Condition ="[Forms]![frmItemInterface]![OrganizationCode]<>\"000\""
    Action ="SetValue"
    Comment ="frmItemInterfaceRouting - Yes"
    Argument ="[Forms]![frmItemInterface]![frmItemInterfaceRouting].[Form].[Visible]"
    Argument ="Yes"
End

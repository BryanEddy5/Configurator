Version =196611
ColumnsShown =0
Begin
    Action ="SetValue"
    Argument ="[OM]"
    Argument ="Null"
End
Begin
    Action ="SetValue"
    Argument ="[BIF]"
    Argument ="Null"
End
Begin
    Condition ="False"
    Action ="SetValue"
    Argument ="[Jacket]"
    Argument ="Null"
End
Begin
    Condition ="False"
    Action ="SetValue"
    Argument ="[Color]"
    Argument ="Null"
End
Begin
    Condition ="False"
    Action ="SetValue"
    Argument ="[ListingCompany]"
    Argument ="Null"
End
Begin
    Condition ="False"
    Action ="SetValue"
    Argument ="[ColorChip]"
    Argument ="Null"
End
Begin
    Condition ="False"
    Action ="Requery"
    Argument ="ListingCompany"
End
Begin
    Action ="Requery"
    Argument ="BIF"
End
Begin
    Action ="Requery"
    Argument ="OM"
End
Begin
    Condition ="False"
    Action ="Requery"
    Argument ="Jacket"
End
Begin
    Condition ="False"
    Action ="Requery"
    Argument ="Color"
End
Begin
    Condition ="False"
    Action ="Requery"
    Argument ="ColorChip"
End
Begin
    Condition ="False"
    Action ="MsgBox"
    Argument ="Customer has a preference to this fiber."
    Argument ="-1"
    Argument ="0"
    Argument ="Fiber Preference"
End
Begin
    Action ="SetValue"
    Argument ="[PrintLine2]"
    Argument ="Null"
End
Begin
    Action ="Requery"
    Argument ="PrintLine2"
End
Begin
    Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
        "nterfaceMacro MinimumClientDesignVersion=\"14.0.0000.0000\" xmlns=\"http://schem"
        "as.microsoft.com/office/accessservices/2009/11/application\"><Statements><Action"
        " Name=\"SetValue\"><Argument Na"
End
Begin
    Comment ="_AXL:me=\"Item\">[OM]</Argument><Argument Name=\"Expression\">Null</Argument></A"
        "ction><Action Name=\"SetValue\"><Argument Name=\"Item\">[BIF]</Argument><Argumen"
        "t Name=\"Expression\">Null</Argument></Action><ConditionalBlock><If><Condition>F"
        "alse</Condition><Statem"
End
Begin
    Comment ="_AXL:ents><Action Name=\"SetValue\"><Argument Name=\"Item\">[Jacket]</Argument><"
        "Argument Name=\"Expression\">Null</Argument></Action></Statements></If></Conditi"
        "onalBlock><ConditionalBlock><If><Condition>False</Condition><Statements><Action "
        "Name=\"SetValue\"><Ar"
End
Begin
    Comment ="_AXL:gument Name=\"Item\">[Color]</Argument><Argument Name=\"Expression\">Null</"
        "Argument></Action></Statements></If></ConditionalBlock><ConditionalBlock><If><Co"
        "ndition>False</Condition><Statements><Action Name=\"SetValue\"><Argument Name=\""
        "Item\">[ListingCompan"
End
Begin
    Comment ="_AXL:y]</Argument><Argument Name=\"Expression\">Null</Argument></Action></Statem"
        "ents></If></ConditionalBlock><ConditionalBlock><If><Condition>False</Condition><"
        "Statements><Action Name=\"SetValue\"><Argument Name=\"Item\">[ColorChip]</Argume"
        "nt><Argument Name=\""
End
Begin
    Comment ="_AXL:Expression\">Null</Argument></Action></Statements></If></ConditionalBlock><"
        "ConditionalBlock><If><Condition>False</Condition><Statements><Action Name=\"Requ"
        "ery\"><Argument Name=\"ControlName\">ListingCompany</Argument></Action></Stateme"
        "nts></If></Conditi"
End
Begin
    Comment ="_AXL:onalBlock><Action Name=\"Requery\"><Argument Name=\"ControlName\">BIF</Argu"
        "ment></Action><Action Name=\"Requery\"><Argument Name=\"ControlName\">OM</Argume"
        "nt></Action><ConditionalBlock><If><Condition>False</Condition><Statements><Actio"
        "n Name=\"Requery\"><Arg"
End
Begin
    Comment ="_AXL:ument Name=\"ControlName\">Jacket</Argument></Action></Statements></If></Co"
        "nditionalBlock><ConditionalBlock><If><Condition>False</Condition><Statements><Ac"
        "tion Name=\"Requery\"><Argument Name=\"ControlName\">Color</Argument></Action></"
        "Statements></If></C"
End
Begin
    Comment ="_AXL:onditionalBlock><ConditionalBlock><If><Condition>False</Condition><Statemen"
        "ts><Action Name=\"Requery\"><Argument Name=\"ControlName\">ColorChip</Argument><"
        "/Action></Statements></If></ConditionalBlock><ConditionalBlock><If><Condition>Fa"
        "lse</Condition><S"
End
Begin
    Comment ="_AXL:tatements><Action Name=\"MessageBox\"><Argument Name=\"Message\">Customer h"
        "as a preference to this fiber.</Argument><Argument Name=\"Title\">Fiber Preferen"
        "ce</Argument></Action></Statements></If></ConditionalBlock><Action Name=\"SetVal"
        "ue\"><Argument Name=\""
End
Begin
    Comment ="_AXL:Item\">[PrintLine2]</Argument><Argument Name=\"Expression\">Null</Argument>"
        "</Action><Action Name=\"Requery\"><Argument Name=\"ControlName\">PrintLine2</Arg"
        "ument></Action></Statements></UserInterfaceMacro>"
End

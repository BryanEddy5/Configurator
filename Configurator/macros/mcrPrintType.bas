Version =196611
ColumnsShown =0
Begin
    Condition ="[Forms]![frmCreateItem]![PrintType]=0"
    Action ="SetValue"
    Argument ="[Customer]"
    Argument ="\"UNPRINTED\""
End
Begin
    Condition ="[Forms]![frmCreateItem]![PrintType]=1"
    Action ="SetValue"
    Argument ="[Customer]"
    Argument ="\"AFL STANDARD\""
End
Begin
    Condition ="[Forms]![frmCreateItem]![PrintType]=3 And [Forms]![frmCreateItem]![PrintLine4] L"
        "ike \"*FEET*\""
    Action ="SetValue"
    Argument ="[Customer]"
    Argument ="\"AFL GENERIC - FEET\""
End
Begin
    Condition ="[Forms]![frmCreateItem]![PrintType]=3 And [Forms]![frmCreateItem]![PrintLine4] N"
        "ot Like \"*FEET*\""
    Action ="SetValue"
    Argument ="[Customer]"
    Argument ="\"AFL GENERIC\""
End
Begin
    Condition ="False"
    Action ="SetValue"
    Argument ="[PrintLine2]"
    Argument ="Null"
End
Begin
    Condition ="[Forms]![frmCreateItem]![PrintType3] Is Null"
    Action ="SetValue"
    Argument ="[PrintLine3]"
    Argument ="Null"
End
Begin
    Condition ="False"
    Action ="SetValue"
    Argument ="[PrintLine4]"
    Argument ="Null"
End
Begin
    Action ="Requery"
    Argument ="PrintLine1"
End
Begin
    Condition ="False"
    Action ="Requery"
    Argument ="PrintLine2"
End
Begin
    Condition ="[Forms]![frmCreateItem]![PrintType3] Is Null"
    Action ="Requery"
    Argument ="PrintLine3"
End
Begin
    Condition ="False"
    Action ="Requery"
    Argument ="PrintLine4"
End
Begin
    Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
        "nterfaceMacro MinimumClientDesignVersion=\"14.0.0000.0000\" xmlns=\"http://schem"
        "as.microsoft.com/office/accessservices/2009/11/application\"><Statements><Condit"
        "ionalBlock><If><Condition>[Fo"
End
Begin
    Comment ="_AXL:rms]![frmCreateItem]![PrintType]=0</Condition><Statements><Action Name=\"Se"
        "tValue\"><Argument Name=\"Item\">[Customer]</Argument><Argument Name=\"Expressio"
        "n\">\"UNPRINTED\"</Argument></Action></Statements></If></ConditionalBlock><Condi"
        "tionalBlock><If><Cond"
End
Begin
    Comment ="_AXL:ition>[Forms]![frmCreateItem]![PrintType]=1</Condition><Statements><Action "
        "Name=\"SetValue\"><Argument Name=\"Item\">[Customer]</Argument><Argument Name=\""
        "Expression\">\"AFL STANDARD\"</Argument></Action></Statements></If></Conditional"
        "Block><ConditionalBlo"
End
Begin
    Comment ="_AXL:ck><If><Condition>[Forms]![frmCreateItem]![PrintType]=3 And [Forms]![frmCre"
        "ateItem]![PrintLine4] Like \"*FEET*\"</Condition><Statements><Action Name=\"SetV"
        "alue\"><Argument Name=\"Item\">[Customer]</Argument><Argument Name=\"Expression\""
        ">\"AFL GENERIC - FEET"
End
Begin
    Comment ="_AXL:\"</Argument></Action></Statements></If></ConditionalBlock><ConditionalBloc"
        "k><If><Condition>[Forms]![frmCreateItem]![PrintType]=3 And [Forms]![frmCreateIte"
        "m]![PrintLine4] Not Like \"*FEET*\"</Condition><Statements><Action Name=\"SetVal"
        "ue\"><Argument Nam"
End
Begin
    Comment ="_AXL:e=\"Item\">[Customer]</Argument><Argument Name=\"Expression\">\"AFL GENERIC"
        "\"</Argument></Action></Statements></If></ConditionalBlock><ConditionalBlock><If"
        "><Condition>False</Condition><Statements><Action Name=\"SetValue\"><Argument Nam"
        "e=\"Item\">[PrintLine2]"
End
Begin
    Comment ="_AXL:</Argument><Argument Name=\"Expression\">Null</Argument></Action></Statemen"
        "ts></If></ConditionalBlock><ConditionalBlock><If><Condition>[Forms]![frmCreateIt"
        "em]![PrintType3] Is Null</Condition><Statements><Action Name=\"SetValue\"><Argum"
        "ent Name=\"Item\">["
End
Begin
    Comment ="_AXL:PrintLine3]</Argument><Argument Name=\"Expression\">Null</Argument></Action"
        "></Statements></If></ConditionalBlock><ConditionalBlock><If><Condition>False</Co"
        "ndition><Statements><Action Name=\"SetValue\"><Argument Name=\"Item\">[PrintLine"
        "4]</Argument><Argum"
End
Begin
    Comment ="_AXL:ent Name=\"Expression\">Null</Argument></Action></Statements></If></Conditi"
        "onalBlock><Action Name=\"Requery\"><Argument Name=\"ControlName\">PrintLine1</Ar"
        "gument></Action><ConditionalBlock><If><Condition>False</Condition><Statements><A"
        "ction Name=\"Requery"
End
Begin
    Comment ="_AXL:\"><Argument Name=\"ControlName\">PrintLine2</Argument></Action></Statement"
        "s></If></ConditionalBlock><ConditionalBlock><If><Condition>[Forms]![frmCreateIte"
        "m]![PrintType3] Is Null</Condition><Statements><Action Name=\"Requery\"><Argumen"
        "t Name=\"ControlNam"
End
Begin
    Comment ="_AXL:e\">PrintLine3</Argument></Action></Statements></If></ConditionalBlock><Con"
        "ditionalBlock><If><Condition>False</Condition><Statements><Action Name=\"Requery"
        "\"><Argument Name=\"ControlName\">PrintLine4</Argument></Action></Statements></I"
        "f></ConditionalBlo"
End
Begin
    Comment ="_AXL:ck></Statements></UserInterfaceMacro>"
End

Operation =1
Option =0
Where ="(((InStr([Expression],'tblDesignCodeFiberElementValue'))<>0)) OR (((InStr([Name1"
    "],'tblDesignCodeFiberElementValue'))<>0)) OR (((InStr([Name2],'tblDesignCodeFibe"
    "rElementValue'))<>0))"
Begin InputTables
    Name ="MSysQueries"
    Name ="MSysObjects"
End
Begin OutputColumns
    Alias ="String"
    Expression ="'tblDesignCodeFiberElementValue'"
    Alias ="QueryName"
    Expression ="MSysObjects.Name"
    Alias ="In Expression"
    Expression ="InStr([Expression],'tblDesignCodeFiberElementValue')"
    Alias ="In Name1"
    Expression ="InStr([Name1],'tblDesignCodeFiberElementValue')"
    Alias ="In Name2"
    Expression ="InStr([Name2],'tblDesignCodeFiberElementValue')"
End
Begin Joins
    LeftTable ="MSysQueries"
    RightTable ="MSysObjects"
    Expression ="MSysQueries.ObjectId = MSysObjects.Id"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xdabddb863458ce4283f7d775a9fd711c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="String"
        dbBinary "GUID" = Begin
            0xdb79edd9531ed44a82a6821986494d65
        End
    End
    Begin
        dbText "Name" ="QueryName"
        dbBinary "GUID" = Begin
            0xa54fa4dd0444894c8fc8e7ce59a9be63
        End
    End
    Begin
        dbText "Name" ="In Expression"
        dbBinary "GUID" = Begin
            0xc9da446601ad6346aaf9aa1faa0775a6
        End
    End
    Begin
        dbText "Name" ="In Name1"
        dbBinary "GUID" = Begin
            0xb240fc2a8f580c44b5a9cb744c3dfa0d
        End
    End
    Begin
        dbText "Name" ="In Name2"
        dbBinary "GUID" = Begin
            0xc6e7674e732f7241a87881442840cef9
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1551
    Bottom =957
    Left =-1
    Top =-1
    Right =1519
    Bottom =633
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MSysQueries"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MSysObjects"
        Name =""
    End
End

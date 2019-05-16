Operation =1
Option =0
Where ="(((vBurnListingBase.ListingCompany)=fORMS!frmCreateItem!ListingCompany) And ((vB"
    "urnListingBase.Base)=Forms!frmCreateItem!Base) And ((vBurnListingBase.US)<>0))"
Begin InputTables
    Name ="vBurnListingBase"
End
Begin OutputColumns
    Expression ="vBurnListingBase.Rating"
    Expression ="vBurnListingBase.ListingConstruction"
    Expression ="vBurnListingBase.ListingSection"
    Expression ="vBurnListingBase.ListingTab"
    Expression ="vBurnListingBase.DocumentID"
    Expression ="vBurnListingBase.ListingCompany"
    Expression ="vBurnListingBase.Base"
    Expression ="vBurnListingBase.US"
    Expression ="vBurnListingBase.Priority"
End
Begin OrderBy
    Expression ="vBurnListingBase.Priority"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4e00fe1d0d3eab4299bc5b33070bd04f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vBurnListingBase.ListingConstruction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.ListingSection"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.ListingTab"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.US"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.Rating"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.ListingCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vBurnListingBase.DocumentID"
        dbLong "AggregateType" ="-1"
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
    Bottom =176
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =97
        Top =7
        Right =320
        Bottom =192
        Top =0
        Name ="vBurnListingBase"
        Name =""
    End
End

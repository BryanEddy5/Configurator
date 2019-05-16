dbMemo "SQL" ="INSERT INTO tblNewBasicCableHold\015\012SELECT [Basic (Cable/Sheath)].*\015\012F"
    "ROM [Basic (Cable/Sheath)] INNER JOIN qryRDBasicCable ON [Basic (Cable/Sheath)]."
    "[New Oracle Part Number]=qryRDBasicCable.[New Oracle Part Number];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBinary "GUID" = Begin
    0xb31175af603a234a99a4bcdeee792ea2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End

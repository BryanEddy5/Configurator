CREATE TABLE [tblOracleItemInterfaceCopied] (
  [OrganizationCode] VARCHAR (3),
  [ItemNumber] VARCHAR (15),
  [ItemCopiedID] AUTOINCREMENT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([OrganizationCode], [ItemNumber])
)

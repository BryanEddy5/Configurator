CREATE TABLE [tblNetworkBOMRoutingsInterface] (
  [OrganizationCode] VARCHAR (3),
  [AssemblyItemNumber] VARCHAR (15),
  [CompletionSubinventory] VARCHAR (10),
  [CommonItemNumber] VARCHAR (15),
  [AlternateRoutingDesignator] VARCHAR (10),
  [ItemCopied] VARCHAR (15),
  [LocationName] VARCHAR (50),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([OrganizationCode], [AssemblyItemNumber])
)

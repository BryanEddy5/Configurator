CREATE TABLE [tblNetworkBOMBillOfMtlsInterface] (
  [OrganizationCode] VARCHAR (3),
  [ItemNumber] VARCHAR (40),
  [AssemblyType] LONG ,
  [AlternateBOMDesignator] VARCHAR (10),
  [CommonOrgCode] VARCHAR (3),
  [CommonItemNumber] VARCHAR (40),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([OrganizationCode], [ItemNumber])
)

CREATE TABLE [tblNetworkBOMOperationSetupParameters] (
  [OrganizationCode] VARCHAR (3),
  [ItemNumber] VARCHAR (15),
  [OperationSeqNum] LONG ,
  [OperationCode] VARCHAR (4),
  [ElementName] VARCHAR (30),
  [ElementValue] VARCHAR (30),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([OrganizationCode], [ItemNumber], [OperationSeqNum], [OperationCode], [ElementName])
)

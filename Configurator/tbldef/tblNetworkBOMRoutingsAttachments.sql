CREATE TABLE [tblNetworkBOMRoutingsAttachments] (
  [OrganizationCode] VARCHAR (3),
  [AssemblyItemNumber] VARCHAR (15),
  [OperationSeqNumber] LONG ,
  [SeqNumber] LONG ,
  [ShortText] LONGTEXT ,
  [Description] VARCHAR (50),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([OrganizationCode], [AssemblyItemNumber], [OperationSeqNumber], [SeqNumber])
)

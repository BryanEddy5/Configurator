CREATE TABLE [tblBOMOpSequencesInterfaceTemp] (
  [OrganizationCode] VARCHAR (3),
  [AssemblyItemNumber] VARCHAR (15),
  [OperationSeqNum] LONG ,
  [AlternateRoutingDesignator] VARCHAR (50),
  [OperationCode] VARCHAR (4),
  [DepartmentCode] VARCHAR (15),
  [EffectivityDate] DATETIME ,
  [OperationDescription] VARCHAR (50),
  [StartupScrap] VARCHAR (150),
  [ShortText] LONGTEXT ,
  [CopyCount] LONG ,
  [ItemCopied] VARCHAR (15),
  [OrgOpSeq] LONG ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([OrganizationCode], [AssemblyItemNumber], [OperationSeqNum], [CopyCount], [OrgOpSeq])
)

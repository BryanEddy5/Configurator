CREATE TABLE [tblNetworkBOMOpSequencesInterface] (
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
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([OrganizationCode], [AssemblyItemNumber], [OperationSeqNum])
)

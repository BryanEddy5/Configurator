CREATE TABLE [tblNetworkBOMOpResourcesInterface] (
  [OrganizationCode] VARCHAR (3),
  [AssemblyItemNumber] VARCHAR (15),
  [AlternateRoutingDesignator] VARCHAR (50),
  [OperationSeqNumber] LONG ,
  [ResourceSeqNumber] LONG ,
  [ResourceCode] VARCHAR (50),
  [EffectivityDate] DATETIME ,
  [StandardRateFlag] VARCHAR (50),
  [AssignedUnits] VARCHAR (50),
  [UsageRateOrAmount] DOUBLE ,
  [UsageRateOrAmountInverse] DOUBLE ,
  [BasisType] VARCHAR (50),
  [ScheduleFlag] VARCHAR (50),
  [AutochargeType] VARCHAR (50),
  [ProcessFlag] VARCHAR (50),
  [TransactionType] VARCHAR (50),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([OrganizationCode], [AssemblyItemNumber], [OperationSeqNumber], [ResourceSeqNumber])
)

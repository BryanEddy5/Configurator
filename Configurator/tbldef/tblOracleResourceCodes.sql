CREATE TABLE [tblOracleResourceCodes] (
  [ResourceCodeID] LONG  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [DepartmentCode] VARCHAR (50),
  [ResourceCode] VARCHAR (20),
  [ResourceCodeDescription] VARCHAR (75),
  [UOM] VARCHAR (5),
  [AllowZeroUsage] BIT 
)

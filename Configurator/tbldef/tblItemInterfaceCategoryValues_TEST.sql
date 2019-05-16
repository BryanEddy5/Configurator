CREATE TABLE [tblItemInterfaceCategoryValues_TEST] (
  [CategorySetName] VARCHAR (50),
  [CategoryValue] VARCHAR (50),
  [DesignTypeID] LONG ,
  [CategoryValueDesc] VARCHAR (255),
  [Qualifier1] VARCHAR (50),
  [Qualifier1Type] VARCHAR (50),
  [Qualifier2] VARCHAR (50),
  [Qualifier2Type] VARCHAR (50),
  [Qualifier1DataType] VARCHAR (50),
  [Qualifier2DataType] VARCHAR (50),
  [Active] BIT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([CategorySetName], [CategoryValue], [DesignTypeID])
)

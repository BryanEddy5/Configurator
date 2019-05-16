CREATE TABLE [tblGetWhereUsedMaterial] (
  [Material] VARCHAR (255),
  [AssemblyItemNumber] VARCHAR (255),
  [Description] VARCHAR (255),
  [MaterialQuantity_PerPrimaryUOM] VARCHAR (255),
  [MaterialQuanitty_PerOrder] DOUBLE ,
  [PrimaryUOM] VARCHAR (255),
  [Item_Status] VARCHAR (255),
  [Customer] VARCHAR (255),
  [OrderNum] VARCHAR (255),
  [Shipped] DATETIME ,
  [Quantity] DOUBLE ,
  [QTY_UOM] VARCHAR (255),
  [CountPerUOM] LONG ,
  [Revenue] DOUBLE ,
  [SoLineNumber] VARCHAR (255),
  [TransactionDate] DATETIME 
)

CREATE TABLE [tblCategories] (
  [CategorySetName] VARCHAR (255),
  [CategoryName] VARCHAR (255),
  [IncludeInReport] BIT ,
  [AppliesTo] LONG ,
  [Active] BIT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([CategorySetName], [CategoryName])
)

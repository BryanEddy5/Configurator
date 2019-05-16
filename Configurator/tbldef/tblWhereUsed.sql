CREATE TABLE [tblWhereUsed] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ComponentItemNumber] VARCHAR (255),
  [AssemblyItemNumber] VARCHAR (255),
  [Quantity] LONG 
)

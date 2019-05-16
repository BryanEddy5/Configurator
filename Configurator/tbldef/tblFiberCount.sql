CREATE TABLE [tblFiberCount] (
  [FiberCount] LONG  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [FiberCountString] VARCHAR (50) CONSTRAINT [FiberCount] UNIQUE 
)

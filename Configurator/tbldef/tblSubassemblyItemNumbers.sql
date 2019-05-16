CREATE TABLE [tblSubassemblyItemNumbers] (
  [SubassemblyItemNumber] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [SubassemblyItemDesc] VARCHAR (255),
  [IncludeInReport] BIT 
)

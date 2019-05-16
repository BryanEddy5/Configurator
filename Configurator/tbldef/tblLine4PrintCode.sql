CREATE TABLE [tblLine4PrintCode] (
  [Line4PrintCode] VARCHAR (50),
  [Line4Print] VARCHAR (100),
  [Spacing] VARCHAR (50),
  [Active] BIT ,
  [Priority] LONG  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Descripton] VARCHAR (255),
  [Field1] VARCHAR (255)
)

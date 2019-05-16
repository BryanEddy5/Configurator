CREATE TABLE [tblSwitchboard] (
  [SwitchboardID] SHORT ,
  [ItemNumber] SHORT ,
  [ItemText] VARCHAR (50),
  [Command] SHORT ,
  [Argument] VARCHAR (50),
  [ID] GUID ,
  [CreatedBy] VARCHAR (50),
  [CreationDate] VARCHAR (50),
   CONSTRAINT [PK_tblSwitchboard] PRIMARY KEY ([SwitchboardID], [ItemNumber])
)

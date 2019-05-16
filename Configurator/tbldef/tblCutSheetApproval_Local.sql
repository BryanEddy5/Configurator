CREATE TABLE [tblCutSheetApproval_Local] (
  [Base] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Oracle] VARCHAR (255),
  [ItemNo] VARCHAR (255),
  [Technical_Approval] BIT ,
  [RequestedDate] DATETIME 
)

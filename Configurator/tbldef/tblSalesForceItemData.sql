CREATE TABLE [tblSalesForceItemData] (
  [Related To] VARCHAR (255),
  [Date] DATETIME ,
  [Start] DATETIME ,
  [Assigned] VARCHAR (255),
  [Priority] VARCHAR (255),
  [Status] VARCHAR (255),
  [Created Date] DATETIME ,
  [Task Waiting Reason] VARCHAR (255),
  [BU] VARCHAR (255),
  [Task Closed Date] DATETIME ,
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Task Subject] VARCHAR (255),
  [TaskAgeBusinessDay] LONG 
)

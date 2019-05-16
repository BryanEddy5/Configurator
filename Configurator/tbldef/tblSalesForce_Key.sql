CREATE TABLE [tblSalesForce_Key] (
  [Related To] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [TaskOpened] DATETIME ,
  [Task Subject] VARCHAR (255),
  [TaskClosed] DATETIME ,
  [Assigned] VARCHAR (255),
  [DaysToCloseTask] DOUBLE ,
  [Date] VARCHAR (255),
  [Task Waiting Reason] VARCHAR (255),
  [CompletedIn48Hours] VARCHAR (255)
)

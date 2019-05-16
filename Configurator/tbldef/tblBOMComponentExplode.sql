CREATE TABLE [tblBOMComponentExplode] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [AssemblyItemNumber] VARCHAR (255),
  [BOMLevel] VARCHAR (255),
  [Weight_kg_km] VARCHAR ,
  [ItemCost_km] DOUBLE ,
  [Cost_Ratio] DOUBLE ,
  [FinishedGood] VARCHAR (255),
  [ComponentItemNumber] VARCHAR (255),
  [Total_Cost_km] DOUBLE 
)

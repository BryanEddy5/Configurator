CREATE TABLE [CutSheetImages_Local] (
  [ImageName] VARCHAR (255),
  [ImageTypeID] VARCHAR (255),
  [Color] VARCHAR (255),
  [Armored] VARCHAR (255),
  [ImageGroupID] VARCHAR (255),
  [ImagePath] VARCHAR (255),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ImageName], [ImageTypeID], [Armored])
)

CREATE TABLE [tblOrders] (
  [OrderID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [CustomerID] LONG ,
  [Item] VARCHAR (50),
  [Qty] LONG ,
  [UnitPrice] CURRENCY 
)

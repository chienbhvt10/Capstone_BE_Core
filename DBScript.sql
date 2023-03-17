use master
GO

DROP DATABASE IF EXISTS attas
GO

CREATE DATABASE attas
ON (NAME = 'attas_data', FILENAME = 'D:\SQL\data\attas_data.mdf') 
LOG ON (NAME = 'attas_log', FILENAME = 'D:\SQL\data\attas_log.ldf');
GO

USE attas
GO
CREATE TABLE [token] (
  [id] int PRIMARY KEY,
  [tokenHash] nvarchar(255),
  [user] nvarchar(255)
)
GO

CREATE TABLE [session] (
  [id] int PRIMARY KEY,
  [sessionHash] nvarchar(255)
)
GO

CREATE TABLE [time] (
  [id] int PRIMARY KEY,
  [sessionId] int,
  [businessId] nvarchar(255)
)
GO

CREATE TABLE [instructor] (
  [id] int PRIMARY KEY,
  [sessionId] int,
  [businessId] nvarchar(255)
)
GO

CREATE TABLE [task] (
  [id] int PRIMARY KEY,
  [sessionId] int,
  [businessId] nvarchar(255)
)
GO

CREATE TABLE [result] (
  [id] int PRIMARY KEY,
  [sessionId] int,
  [tokenId] int,
  [taskId] int,
  [instructorId] int,
  [timeId] int
)
GO

ALTER TABLE [time] ADD FOREIGN KEY ([sessionId]) REFERENCES [session] ([id])
GO

ALTER TABLE [instructor] ADD FOREIGN KEY ([sessionId]) REFERENCES [session] ([id])
GO

ALTER TABLE [task] ADD FOREIGN KEY ([sessionId]) REFERENCES [session] ([id])
GO

ALTER TABLE [result] ADD FOREIGN KEY ([sessionId]) REFERENCES [session] ([id])
GO

ALTER TABLE [result] ADD FOREIGN KEY ([taskId]) REFERENCES [task] ([id])
GO

ALTER TABLE [result] ADD FOREIGN KEY ([instructorId]) REFERENCES [instructor] ([id])
GO

ALTER TABLE [result] ADD FOREIGN KEY ([timeId]) REFERENCES [time] ([id])
GO

ALTER TABLE [result] ADD FOREIGN KEY ([tokenId]) REFERENCES [token] ([id])
GO

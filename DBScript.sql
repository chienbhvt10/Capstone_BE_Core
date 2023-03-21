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
  [id] int IDENTITY(1,1) PRIMARY KEY,
  [tokenHash] nvarchar(255),
  [user] nvarchar(255)
)
GO

CREATE TABLE [session] (
  [id] int IDENTITY(1,1) PRIMARY KEY,
  [sessionHash] nvarchar(255),
  [statusId] int,
  [solutionCount] int
)
GO

CREATE TABLE [status] (
  [id] int IDENTITY(1,1) PRIMARY KEY,
  [name] nvarchar(255)
)
GO

CREATE TABLE [time] (
  [id] int IDENTITY(1,1) PRIMARY KEY,
  [sessionId] int,
  [businessId] nvarchar(255),
  [order] int
)
GO

CREATE TABLE [instructor] (
  [id] int IDENTITY(1,1) PRIMARY KEY,
  [sessionId] int,
  [businessId] nvarchar(255),
  [order] int
)
GO

CREATE TABLE [task] (
  [id] int IDENTITY(1,1) PRIMARY KEY,
  [sessionId] int,
  [businessId] nvarchar(255),
  [order] int
)
GO

CREATE TABLE [result] (
  [id] int IDENTITY(1,1) PRIMARY KEY,
  [solutionId] int,
  [taskOrder] int,
  [instructorOrder] int,
  [timeOrder] int
)
GO

CREATE TABLE [solution] (
  [id] int IDENTITY(1,1) PRIMARY KEY,
  [sessionId] int,
  [no] int,
  [taskAssigned] int,
  [slotCompability] int,
  [subjectDiversity] int,
  [quotaAvailable] int,
  [walkingDistance] int,
  [subjectPreference] int,
  [slotPreference] int
)
GO

ALTER TABLE [time] ADD FOREIGN KEY ([sessionId]) REFERENCES [session] ([id])
GO

ALTER TABLE [instructor] ADD FOREIGN KEY ([sessionId]) REFERENCES [session] ([id])
GO

ALTER TABLE [task] ADD FOREIGN KEY ([sessionId]) REFERENCES [session] ([id])
GO

ALTER TABLE [result] ADD FOREIGN KEY ([solutionId]) REFERENCES [solution] ([id])
GO

ALTER TABLE [session] ADD FOREIGN KEY ([statusId]) REFERENCES [status] ([id])
GO

ALTER TABLE [solution] ADD FOREIGN KEY ([sessionId]) REFERENCES [session] ([id])
GO

INSERT INTO token (tokenHash,[user]) VALUES ('token','FPT');

INSERT INTO [status] (name) VALUES ('PENDING')
INSERT INTO [status] (name) VALUES ('UNKNOWN')
INSERT INTO [status] (name) VALUES ('INFEASIBLE')
INSERT INTO [status] (name) VALUES ('FEASIBLE')
INSERT INTO [status] (name) VALUES ('OPTIMAL')


select * from [session]
select * from solution
select * from result
select * from task
select * from [time]
select * from instructor
select * from [status]
select * from token
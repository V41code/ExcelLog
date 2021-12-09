USE [SPData]
GO

/****** Object:  Table [dbo].[ExlLogFiles]    Script Date: 09.12.2021 12:25:33 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ExlLogFiles](
	[FileId] [int] IDENTITY(1,1) NOT NULL,
	[FileName] [nchar](255) NULL,
	[CrUserName] [nchar](100) NULL,
	[CrDt] [datetime] NULL
) ON [PRIMARY]

GO

EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Имя файл, который протоколируется' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ExlLogFiles', @level2type=N'COLUMN',@level2name=N'FileName'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Кто создал запись о файле ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ExlLogFiles', @level2type=N'COLUMN',@level2name=N'CrUserName'
GO

USE [SPData]
GO

/****** Object:  Table [dbo].[ExlLogSheets]    Script Date: 09.12.2021 12:26:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ExlLogSheets](
	[SheetId] [int] IDENTITY(1,1) NOT NULL,
	[FK_ExlLogFileId] [int] NOT NULL,
	[SheetName] [nchar](255) NOT NULL,
	[CrUserName] [nchar](100) NOT NULL,
	[CrDt] [datetime] NOT NULL
) ON [PRIMARY]

GO

EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'Лист файла' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'ExlLogSheets', @level2type=N'COLUMN',@level2name=N'SheetName'
GO

USE [SPData]
GO

/****** Object:  Table [dbo].[ExlLogData]    Script Date: 09.12.2021 12:26:32 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ExlLogData](
	[Rid] [int] IDENTITY(1,1) NOT NULL,
	[FK_FileId] [int] NULL,
	[FK_SheetId] [int] NOT NULL,
	[RowIndex] [int] NOT NULL,
	[ColIndex] [int] NULL,
	[NewValue] [nvarchar](max) NULL,
	[UserName] [nvarchar](100) NULL,
	[CrDt] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO


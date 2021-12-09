USE [SPData]
GO
/****** Object:  StoredProcedure [dbo].[uspExlLogDataInsert]    Script Date: 09.12.2021 12:06:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Val
-- Create date: 20.09.21
-- Description:	Записать в лог Ид листа, строку, колонку, имя пользователя и логируемое значение
-- =============================================
--CREATE PROCEDURE uspExlLogDataInsert 
ALTER PROCEDURE [dbo].[uspExlLogDataInsert]
	-- parameters for the stored procedure here
	@res numeric(18,0) output,
	@SheetId int,
	@RowIndex int,
	@ColIndex int,
	@NewValue nvarchar(4000)
AS
BEGIN
	DECLARE @TmpTable TABLE (Id INT,Name nvarchar(255));
	DECLARE @lFileId Int
	DECLARE @lSheetId int
	DECLARE @sUserName nvarchar(100)
	DECLARE @CrDt DateTime=GETDATE()
	--DECLARE @lRes Int
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
	SELECT @sUserName=SYSTEM_USER;
	SELECT @lFileId=FK_ExlLogFileId FROM [dbo].[ExlLogSheets] WHERE SheetId=@SheetId;

	
	INSERT INTO ExlLogData([FK_FileId],[FK_SheetId],[RowIndex],[ColIndex],[NewValue],[UserName],[CrDt]) 
	--OUTPUT INSERTED.Rid INTO @res --SheetId,INSERTED.SheetName INTO @TmpTable
		VALUES(@lFileId,@SheetId,@RowIndex,@ColIndex,@NewValue,@sUserName,@crDt)
	
	select @res=IDENT_CURRENT('ExlLogData')
END

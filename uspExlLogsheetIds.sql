USE [SPData]
GO
/****** Object:  StoredProcedure [dbo].[uspExlLogGetFileSheetIds]    Script Date: 09.12.2021 12:07:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Val
-- Create date: 18.09.21
-- Description:	Выдать Id файла и Иды листов
-- =============================================
--CREATE PROCEDURE uspExlLogGetFileSheetIds 
ALTER PROCEDURE [dbo].[uspExlLogGetFileSheetIds] 
	-- parameters for the stored procedure here
	@FileName nvarchar(255),
	@SheetNameList nvarchar(2000)=null
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
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
	SELECT @lFileId=FileId FROM [dbo].[ExlLogFiles] WHERE FileName=@FileName;
	IF ISNULL(@lFileId,-1)<1 
		BEGIN
			INSERT INTO ExlLogFiles (FileName,CrUSerName,CrDt) 
			OUTPUT INSERTED.FileId,INSERTED.FileName INTO @TmpTable --VALUES(Id,Name)
			VALUES(@FileName,@sUserName,@CrDt)

			SELECT @lFileId=(SELECT max(Id) FROM @TmpTable)
		END
	INSERT INTO ExlLogSheets (FK_ExlLogFileId,SheetName,CrUserNAme,CrDt) 
	--OUTPUT INSERTED.* --SheetId,INSERTED.SheetName INTO @TmpTable
	SELECT @lFileId as FK_ExlLogFileId,list.Value as SheetName,@sUserName,@CrDt
		FROM STRING_SPLITN(@SheetNameList,',') as list
		WHERE NOT EXISTS(SELECT SheetId FROM ExlLogSheets sh WHERE sh.FK_ExlLogFileId=@lFileId  AND  list.Value=sh.SheetName)
	;
--	SELECT * FROM STRING_SPLITN(@SheetNameList,',')
	--SELECT * FROM @TmpTable
	
	SELECT * FROM ExlLogSheets WHERE FK_ExlLogFileId=@lFileId;
END

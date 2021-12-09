# ExcelLog
Excel cells change log. VBA module in workbook write changes to SQL Server database

handle events for WorkBook Open,Close,Save
handle event for sheet Change
tables on server
  ExlLogFiles - list of files to log - adds, when first log action performed
      - FileId
      - FileName
  ExlLogSheets - list of sheets to log - adds, when first log action performed
      - SheetId
      - FK_FileId
      - SheetName
  ExlLogData - log data
      - Rid - record identity
      - FK_FileId
      - FK_SheetId
      - RowIndex
      - ColIndex  - cell addres (row and column) in Sheet
      - NewValue - string
      - UserName - system login name (Environment)
      - Log DateTime

/****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP (1000) [Company]
      ,[Strasse_Hausnummer_Postfach]
      ,[Postleitzahl]
      ,[Stadt]
      ,[Land]
      ,[keyAddress]
  FROM [CAD].[dbo].[tAC_Addresses] order by LEN(keyAddress) DESC, keyAddress DESC

  SELECT TOP (1000) [Company]
      ,[OrderNo]
      ,[tsScreenshotCreated]
      ,[idxAddress]
  FROM [CAD].[dbo].[tAC_ProdScreenshots] 

SELECT TOP (1000) [Company]
      ,[keyAddress]
      ,[tsScreenshotCreated]
      ,[locationPath]
      ,[pngFileName]
      ,[register]
FROM [CAD].[dbo].[tAC_Screenshots] ORDER BY tsScreenshotCreated DESC

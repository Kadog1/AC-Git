SELECT [Company]
      ,[Strasse_Hausnummer_Postfach]
      ,[Postleitzahl]
      ,[Stadt]
      ,[Land]
  FROM [CAD].[dbo].[tAC_Addresses_TEST]

SELECT [Company]
      ,[tsScreenshotCreated]
      ,[locationPath]
      ,[pngFileName]
  FROM [CAD].[dbo].[tAC_Screenshots_TEST]

  SELECT [Company]
      ,[OrderNo]
      ,[idxAddress]
      ,[tsScreenshotCreated]
  FROM [CAD].[dbo].[tAC_ProdScreenshots_TEST]

--DELETE FROM [CAD].[dbo].[tAC_Screenshots_TEST]
--DELETE FROM [CAD].[dbo].[tAC_ProdScreenshots_TEST]
--DELETE FROM [CAD].[dbo].[tAC_Addresses_TEST]

---
---INSERT INTO [CAD].[dbo].[tAC_Addresses_TEST] (Company, Strasse_Hausnummer_Postfach, Postleitzahl, Stadt, Land)
---SELECT TOP (1) 'Lufthansa AirPlus Servicekarten GmbH', 'Dornhofstraﬂe 10', '63263', 'Neu-Isenburg', 'DE'
 ---FROM [CAD].[dbo].[tAC_Addresses_TEST]
 ---WHERE NOT EXISTS (SELECT * FROM [CAD].[dbo].[tAC_Addresses_TEST]
 ---WHERE Company = 'Lufthansa AirPlus Servicekarten GmbH' AND Strasse_Hausnummer_Postfach = 'Dornhofstraﬂe 10' AND Postleitzahl = '63263' AND Stadt = 'Neu-Isenburg' AND Land = 'DE')
 ---






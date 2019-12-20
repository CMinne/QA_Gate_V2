DECLARE @var_Keyence INT

IF NOT EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'[dbo].[Keyence_Data_Test]')
AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
CREATE TABLE [dbo].[Keyence_Data_Test] (
    [Id_Keyence]      INT IDENTITY (1, 1) NOT NULL,
    [Double_Taillage] BIT,
    [Coup_Denture_1]  BIT,
    [Coup_Denture_2]  BIT,
    [Chanfrein_1]     BIT,
    [Chanfrein_2]     BIT,
    [Chanfrein_3]     BIT,
    [Chanfrein_4]     BIT,
	[Heure_Reseau]    DATETIME DEFAULT GETDATE()
    PRIMARY KEY CLUSTERED ([Id_Keyence] ASC)
);

CREATE TABLE [dbo].[#tempData] (
    [Id_Keyence]      INT IDENTITY (1, 1) NOT NULL,
    [Double_Taillage] BIT,
    [Coup_Denture_1]  BIT,
    [Coup_Denture_2]  BIT,
    [Chanfrein_1]     BIT,
    [Chanfrein_2]     BIT,
    [Chanfrein_3]     BIT,
    [Chanfrein_4]     BIT,
    PRIMARY KEY CLUSTERED ([Id_Keyence] ASC)
);

BULK INSERT #tempDATA
FROM 'C:\Users\Public\Documents\KEYENCE\CV-H1X\QA Gate 4.0\SD2\cv-x\result\SD1_008\191010_123145.csv'
WITH
(
FIRSTROW = 1,
FIELDTERMINATOR = ',',
ROWTERMINATOR = '\r'
--ERRORFILE = 'C:\Users\chamin\source\repos\QA_GATE\logfile.log'
)

SELECT @var_Keyence = max(Id_Keyence) FROM #tempData
INSERT INTO Keyence_Data_Test (Double_Taillage, Coup_Denture_1, Coup_Denture_2, Chanfrein_1, Chanfrein_2, Chanfrein_3, Chanfrein_4)
SELECT Double_Taillage, Coup_Denture_1, Coup_Denture_2, Chanfrein_1, Chanfrein_2, Chanfrein_3, Chanfrein_4 FROM #tempData WHERE Id_Keyence=(@var_Keyence-3)
SELECT * FROM Keyence_Data_Test
DROP Table #tempData



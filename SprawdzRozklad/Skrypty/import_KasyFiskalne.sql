USE [%DB%]
GO

SET NOCOUNT ON;
--tworzenie tabeli
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[IMP_KasyFiskalne]') AND type in (N'U'))
DROP TABLE [dbo].[IMP_KasyFiskalne]
GO

CREATE TABLE IMP_KasyFiskalne
(
ID varchar(max),
LOGO varchar(max),
NRINW varchar(max),
NRFABR varchar(max),
NREWUS varchar(max),
DATAPROD varchar(max),
TYPBIL varchar(max),
DATAPRZ varchar(max),
NRPRF varchar(max),
DATAOPRF varchar(max),
NAZRFZPRF varchar(max),
LRZPRF varchar(max),
DATAZPRF varchar(max),
NAZRZZPRF varchar(max),
NRRZZPRF varchar(max),
NRKPPRF varchar(max),
NRPRZPRF varchar(max),
IMIEZPRF varchar(max),
NAZWZPRF varchar(max),
STATZPRF varchar(max),
NRORF varchar(max),
DATAOORF varchar(max),
NAZRFZORF varchar(max),
LRZORF varchar(max),
NAZRZORF varchar(max),
NRRZORF varchar(max),
DATAZORF varchar(max),
NRKPORF varchar(max),
NRPRZORF varchar(max),
IMIEZORF varchar(max),
NAZWZORF varchar(max),
STATZORF varchar(max),
LRF varchar(max),
DATAWBF varchar(max),
LOGOPOWBF varchar(max),
DATAKON varchar(max),
DATAFISK varchar(max),
DATAPRZEGL varchar(max),
ADRES varchar(max),
KODP varchar(max),
MIEJSC varchar(max),
NRDECUS varchar(max),
DATADECUS varchar(max),
DATAZAK varchar(max),
LMIESGWAR varchar(max),
UWAGI varchar(max),
BEZOPSERW varchar(max),
WERPSB varchar(max),
DATAOP varchar(max),
GODZOP varchar(max),
NR_SLUZBOP varchar(max)
);

--START

--KONIEC
GO


declare @MyCount int
select @MyCount=Count(*) from IMP_KasyFiskalne
PRINT 'Wprowadzono kas fiskalnych do tabeli tymczasowej: '+Str(@MyCount)

--TUTAJ: START POPRAWEK DANYCH miejsce na kod do uporzadkowania danych: usuniêcie niepotrzebnych, poprawki, itp.

DELETE FROM dbo.IMP_KasyFiskalne WHERE LOGO IS NULL --usuwam bzdurne rekordy
PRINT 'Usuniête kasy z pustym LOGO: '+Str(@@ROWCOUNT)

GO

--usuniêcie starych (data ostatniego RF starsza ni¿ 1 rok)
--SELECT GETDATE()-365, DATAOORF,* FROM dbo.IMP_KasyFiskalne WHERE convert(DATETIME,DATAOORF,104)< GETDATE()-365
DELETE FROM dbo.IMP_KasyFiskalne WHERE convert(DATETIME,DATAOORF,104)< GETDATE()-365
PRINT 'Usuniêto kas fiskanych DATAOORF starsza ni¿ rok: '+Str(@@ROWCOUNT)
declare @MyCount int
select @MyCount=Count(*) from IMP_KasyFiskalne
PRINT 'Kas fiskalnych w tabeli tymczasowej: '+Str(@MyCount)
GO

IF EXISTS(SELECT LOGO, COUNT(1) FROM dbo.IMP_KasyFiskalne GROUP BY LOGO, TYPBIL HAVING COUNT(1)>1)
BEGIN  
	PRINT 'S¥ REKORDY Z DUPLIKATAMI LOGO. Import przerwany'  
	SELECT LOGO as 'Duplikaty LOGO', COUNT(1) AS ILOSC FROM dbo.IMP_KasyFiskalne GROUP BY LOGO, TYPBIL HAVING COUNT(1)>1

	--NAPRAWA DUPLIKATOW LOGO
	--podwójne numery LOGO: nie mo¿e byæ! LOGO u¿ywam jako klucz przy imporcie
	--SELECT * FROM dbo.IMP_KasyFiskalne WHERE LOGO='ABC 123'
	--DELETE FROM dbo.IMP_KasyFiskalne WHERE LOGO='???' AND NRINW='???'
	--SELECT * FROM dbo.IMP_KasyFiskalne
END

--KONIEC!!!! POPRAWEK DANYCH


--import danych - MAT_TicketRegister
--SELECT * FROM dbo.MAT_TicketRegister
--SELECT * FROM dbo.MAT_TicketRegisterHistory
--SELECT * FROM dbo.IMP_KasyFiskalne

ELSE BEGIN
	INSERT INTO dbo.MAT_TicketRegister
	( CurrentLogo ,CurrentFiscalizationDate ,TRSN ,TaxNumber ,FirmwareVersion ,InService ,InServiceStartDate ,ScrapDate ,StockNumber ,
	ProductionDate ,PurchaseDate ,Guarantee ,DecisionDate ,DecisionNumber ,Company_ID ,Description 
	,Type ,TicketLetter ,TicketCode ,
	CREATED ,MODIFIED ,BusPC_ID ,CompanyOwner_ID ,CompanySeller_ID ,NextMaintenanceDate)
	SELECT LOGO, 
	ISNULL(convert(DATETIME,ISNULL(ISNULL(DATAFISK,DATAPRZ),DATAOPRF),104),'2017-01-01')
	, NRFABR, NREWUS, WERPSB, 1, 
	ISNULL(convert(DATETIME,DATAPRZ,104),'2017-01-01'), 
	NULL,NRINW,
	convert(DATETIME,DATAPROD,104), 
	convert(DATETIME,DATAZAK,104), 
	LMIESGWAR,
	convert(DATETIME,DATADECUS,104),NRDECUS, NULL,UWAGI
	, CASE TYPBIL 
	    WHEN 'EMAR-105' THEN 2 
		WHEN 'EMAR-205' THEN 4
		WHEN 'PRINTO LINE B' THEN 1 
		WHEN 'DUO BUS' THEN 1 
		WHEN 'POSNET' THEN 1 
		WHEN 'NOVITUS' THEN 1 
		WHEN 'PRINTO BUS' THEN 1 ELSE 0 END, '', '',
	GETDATE(), GETDATE(), NULL, 1, NULL, 
	convert(DATETIME,DATAPRZEGL,104)	
	FROM 
	dbo.IMP_KasyFiskalne k
	LEFT JOIN dbo.MAT_TicketRegister t ON k.LOGO=t.CurrentLogo
	WHERE 
	t. ID IS NULL

	

	PRINT 'WPROWADZONO kas fiskanych do INFORMICA: '+Str(@@ROWCOUNT)
	declare @MyCount int
	select @MyCount=Count(*) from dbo.MAT_TicketRegister where Type=1
	PRINT 'W tym:'
    PRINT '- PRINTO BUS / DUO BUS / PRINTO LINE B / POSNET / NOVITUS: '+Str(@MyCount)
	select @MyCount=Count(*) from dbo.MAT_TicketRegister where Type=2
    PRINT '- EMAR-105: '+Str(@MyCount)
	select @MyCount=Count(*) from dbo.MAT_TicketRegister where Type=4
    PRINT '- EMAR-205: '+Str(@MyCount)
	select @MyCount=Count(*) from dbo.MAT_TicketRegister where Type=0
	if @MyCount>0
    PRINT '- TYP NUMER 0: '+Str(@MyCount)+' DO SPRAWDZENIA !!!!!!!!!!!!!!'


	INSERT INTO dbo.MAT_TicketRegisterHistory
	( TicketRegister_ID ,FiscalizationDate ,Logo ,MaintenanceDate ,
	RepairDate ,Description ,CREATED ,MODIFIED)
	SELECT t.ID,t.CurrentFiscalizationDate,t.CurrentLogo, t.NextMaintenanceDate,
	NULL, 'import z PP/KK', GETDATE(), GETDATE()
	FROM 
	dbo.IMP_KasyFiskalne k
	INNER JOIN dbo.MAT_TicketRegister t
	ON k.LOGO=t.CurrentLogo
	LEFT JOIN dbo.MAT_TicketRegisterHistory h
	ON t.ID = h.TicketRegister_ID
	WHERE 
	h.ID IS NULL

	--PRINT 'Wprowadzono rekordów zdarzeñ dla kas fiskanych do INFORMICA: '+Str(@@ROWCOUNT)
END --else

GO
DROP TABLE dbo.IMP_KasyFiskalne;
GO


--SELECT * FROM dbo.MAT_Bus
--SELECT * FROM dbo.MAT_BusTypeB
--UPDATE dbo.mat_bus SET BusTypeB_ID=2 WHERE ID>10
--SELECT * FROM dbo.PLAN_Driver WHERE DriverType_ID=1
--SELECT * FROM dbo.PLAN_DriverType
--UPDATE dbo.PLAN_Driver SET DriverType_ID=100 WHERE DriverType_ID=1 AND ID>21
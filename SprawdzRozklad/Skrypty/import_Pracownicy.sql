USE [master]
GO
IF left(cast(SERVERPROPERTY('productversion') as varchar),2)>=13
ALTER DATABASE [%DB%] SET COMPATIBILITY_LEVEL = 130 --przestawiam wersjê na 2016
GO

USE [%DB%]
GO

--tworzenie tabeli
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[IMP_pracownicy]') AND type in (N'U'))
DROP TABLE [dbo].[IMP_pracownicy]
GO

CREATE TABLE IMP_pracownicy
(
NR_SLUZB varchar(max),
IMIE varchar(max),
NAZWISKO varchar(max),
STATUS varchar(max),
ODDZIAL varchar(max),
DATA_PRZYJ varchar(max),
DATA_ZWOL varchar(max),
TRYBROZWUM varchar(max),
ADRES varchar(max),
KODP varchar(max),
MIEJSC varchar(max),
SYM varchar(max),
GMINA varchar(max),
NAZKR varchar(max),
KOD2 varchar(max),
KIER varchar(max),
TEL varchar(max),
TELKOM varchar(max),
ROKROZL varchar(max),
PLANDNIURL varchar(max),
DNIURLDO varchar(max),
DATAURL1 varchar(max),
LBDNIROB1 varchar(max),
DATAURL2 varchar(max),
LBDNIROB2 varchar(max),
DATAURL3 varchar(max),
LBDNIROB3 varchar(max),
NRREJAUT varchar(max),
DATPRZYAUT varchar(max),
NRINWBIL varchar(max),
LOGOPB varchar(max),
DATPRZYBIL varchar(max),
NRKARPAM varchar(max),
DATPRZKP varchar(max),
NAZPRZ varchar(max),
NRPRZ varchar(max),
DATAREJP varchar(max),
GODZREJP varchar(max),
DATAPOCZP varchar(max),
DATAZAKP varchar(max),
NRPRF varchar(max),
LOGOPRF varchar(max),
NRPZAD varchar(max),
NRPKARDR varchar(max),
NRKPPZAD varchar(max),
NAZORZ varchar(max),
NRORZ varchar(max),
DATAREJO varchar(max),
GODZREJO varchar(max),
DATAPOCZO varchar(max),
DATAZAKO varchar(max),
NRORF varchar(max),
LOGOORF varchar(max),
NRZAD varchar(max),
NRKARDR varchar(max),
NRKPOZAD varchar(max),
NRSPLAC varchar(max),
WYSTNOTY varchar(max),
WYSTFAKT varchar(max),
DRUKDOK varchar(max),
ETAT varchar(max),
NUMERPRJ varchar(max),
DATAWYDPRJ varchar(max),
DATAWAZPRJ varchar(max),
PRAWOJ varchar(max),
DWAZLEK varchar(max),
DWAZPSYCH varchar(max),
DATAPKRT varchar(max),
DATAKKRT varchar(max),
NUMERKRT varchar(max),
WYDKRT varchar(max),
DATA_UR varchar(max),
NR_DOWOD varchar(max),
NR_PASZ varchar(max),
PESEL varchar(max),
DATABLEK varchar(max),
DATABPSYCH varchar(max),
NRZASKDPO varchar(max),
DATKDPO varchar(max),
NRZASKDPR varchar(max),
DATKDPR varchar(max),
NRTELKOM varchar(max),
DTELKOM varchar(max),
DATAOP varchar(max),
GODZOP varchar(max),
NR_SLUZBOP varchar(max)
);

--START

--KONIEC
GO

DELETE FROM dbo.IMP_pracownicy 
WHERE (STATUS NOT IN ('7','8')) --zostawiam tylko kierowców i kasjerów
GO


--TUTAJ: START POPRAWEK DANYCH miejsce na kod do uporzadkowania danych: usuniêcie niepotrzebnych, poprawki liczby miejsc, itp.

--SELECT * FROM dbo.IMP_Pracownicy

--podwójne numery s³u¿bowe kierowców: nie mo¿e byæ! numer s³u¿bowy u¿ywam jako klucz przy imporcie
--SELECT nr_sluzb, COUNT(1) FROM dbo.IMP_Pracownicy GROUP BY nr_sluzb HAVING COUNT(1)>1
--SELECT * FROM dbo.IMP_Pracownicy WHERE nr_sluzb='572FL'
--DELETE FROM dbo.IMP_Pracownicy WHERE nr_sluzb='???' AND NAZWISKO='???'

--usuniêcie niepracuj¹cych
DELETE FROM dbo.IMP_Pracownicy WHERE DATA_ZWOL IS NOT NULL 

--poprawienie formatu dat
--SELECT isnull(isnull(isnull(isnull(TRY_CONVERT(DATE,DATA_PRZYJ,104),TRY_CONVERT(DATE,DATA_PRZYJ,103)),TRY_CONVERT(DATE,DATA_PRZYJ,102)),TRY_CONVERT(DATE,DATA_PRZYJ,101)),TRY_CONVERT(DATE,DATA_PRZYJ,101)), * FROM dbo.IMP_Pracownicy
-- This script updates the date columns by attempting to parse several date formats.
-- It is written to be compatible with SQL Server 2008 R2, which does not have TRY_CONVERT.

UPDATE dbo.IMP_Pracownicy
SET DATA_PRZYJ =
    CASE
        -- Attempt to convert from format 104 (dd.mm.yyyy)
        WHEN DATA_PRZYJ LIKE '[0-3][0-9].[0-1][0-9].[12][0-9][0-9][0-9]' AND ISDATE(SUBSTRING(DATA_PRZYJ,7,4)+'-'+SUBSTRING(DATA_PRZYJ,4,2)+'-'+SUBSTRING(DATA_PRZYJ,1,2)) = 1
            THEN CONVERT(DATE, DATA_PRZYJ, 104)
        -- Attempt to convert from format 103 (dd/mm/yyyy)
        WHEN DATA_PRZYJ LIKE '[0-3][0-9]/[0-1][0-9]/[12][0-9][0-9][0-9]' AND ISDATE(SUBSTRING(DATA_PRZYJ,7,4)+'-'+SUBSTRING(DATA_PRZYJ,4,2)+'-'+SUBSTRING(DATA_PRZYJ,1,2)) = 1
            THEN CONVERT(DATE, DATA_PRZYJ, 103)
        -- Attempt to convert from format 102 (yy.mm.dd)
        WHEN DATA_PRZYJ LIKE '[0-9][0-9].[0-1][0-9].[0-3][0-9]' AND ISDATE('20' + SUBSTRING(DATA_PRZYJ,1,2) + '-' + SUBSTRING(DATA_PRZYJ,4,2) + '-' + SUBSTRING(DATA_PRZYJ,7,2)) = 1
            THEN CONVERT(DATE, DATA_PRZYJ, 102)
        -- Attempt to convert from format 101 (mm/dd/yy)
        WHEN DATA_PRZYJ LIKE '[0-1][0-9]/[0-3][0-9]/[0-9][0-9]' AND ISDATE('20' + SUBSTRING(DATA_PRZYJ,7,2) + '-' + SUBSTRING(DATA_PRZYJ,1,2) + '-' + SUBSTRING(DATA_PRZYJ,4,2)) = 1
            THEN CONVERT(DATE, DATA_PRZYJ, 101)
        ELSE NULL
    END;

UPDATE dbo.IMP_Pracownicy
SET DATA_ZWOL =
    CASE
        -- Attempt to convert from format 104 (dd.mm.yyyy)
        WHEN DATA_ZWOL LIKE '[0-3][0-9].[0-1][0-9].[12][0-9][0-9][0-9]' AND ISDATE(SUBSTRING(DATA_ZWOL,7,4)+'-'+SUBSTRING(DATA_ZWOL,4,2)+'-'+SUBSTRING(DATA_ZWOL,1,2)) = 1
            THEN CONVERT(DATE, DATA_ZWOL, 104)
        -- Attempt to convert from format 103 (dd/mm/yyyy)
        WHEN DATA_ZWOL LIKE '[0-3][0-9]/[0-1][0-9]/[12][0-9][0-9][0-9]' AND ISDATE(SUBSTRING(DATA_ZWOL,7,4)+'-'+SUBSTRING(DATA_ZWOL,4,2)+'-'+SUBSTRING(DATA_ZWOL,1,2)) = 1
            THEN CONVERT(DATE, DATA_ZWOL, 103)
        -- Attempt to convert from format 102 (yy.mm.dd)
        WHEN DATA_ZWOL LIKE '[0-9][0-9].[0-1][0-9].[0-3][0-9]' AND ISDATE('20' + SUBSTRING(DATA_ZWOL,1,2) + '-' + SUBSTRING(DATA_ZWOL,4,2) + '-' + SUBSTRING(DATA_ZWOL,7,2)) = 1
            THEN CONVERT(DATE, DATA_ZWOL, 102)
        -- Attempt to convert from format 101 (mm/dd/yy)
        WHEN DATA_ZWOL LIKE '[0-1][0-9]/[0-3][0-9]/[0-9][0-9]' AND ISDATE('20' + SUBSTRING(DATA_ZWOL,7,2) + '-' + SUBSTRING(DATA_ZWOL,1,2) + '-' + SUBSTRING(DATA_ZWOL,4,2)) = 1
            THEN CONVERT(DATE, DATA_ZWOL, 101)
        ELSE NULL
    END;

UPDATE dbo.IMP_Pracownicy
SET DATA_UR =
    CASE
        -- Attempt to convert from format 104 (dd.mm.yyyy)
        WHEN DATA_UR LIKE '[0-3][0-9].[0-1][0-9].[12][0-9][0-9][0-9]' AND ISDATE(SUBSTRING(DATA_UR,7,4)+'-'+SUBSTRING(DATA_UR,4,2)+'-'+SUBSTRING(DATA_UR,1,2)) = 1
            THEN CONVERT(DATE, DATA_UR, 104)
        -- Attempt to convert from format 103 (dd/mm/yyyy)
        WHEN DATA_UR LIKE '[0-3][0-9]/[0-1][0-9]/[12][0-9][0-9][0-9]' AND ISDATE(SUBSTRING(DATA_UR,7,4)+'-'+SUBSTRING(DATA_UR,4,2)+'-'+SUBSTRING(DATA_UR,1,2)) = 1
            THEN CONVERT(DATE, DATA_UR, 103)
        -- Attempt to convert from format 102 (yy.mm.dd)
        WHEN DATA_UR LIKE '[0-9][0-9].[0-1][0-9].[0-3][0-9]' AND ISDATE('20' + SUBSTRING(DATA_UR,1,2) + '-' + SUBSTRING(DATA_UR,4,2) + '-' + SUBSTRING(DATA_UR,7,2)) = 1
            THEN CONVERT(DATE, DATA_UR, 102)
        -- Attempt to convert from format 101 (mm/dd/yy)
        WHEN DATA_UR LIKE '[0-1][0-9]/[0-3][0-9]/[0-9][0-9]' AND ISDATE('20' + SUBSTRING(DATA_UR,7,2) + '-' + SUBSTRING(DATA_UR,1,2) + '-' + SUBSTRING(DATA_UR,4,2)) = 1
            THEN CONVERT(DATE, DATA_UR, 101)
        ELSE NULL
    END;



--KONIEC!!!! POPRAWEK DANYCH


--import danych - PLAN_Driver
--SELECT * FROM dbo.PLAN_Driver
--SELECT * FROM dbo.PLAN_DriverType
--SELECT * FROM dbo.PLAN_DriverActionType
--SELECT * FROM dbo.PLAN_DriverAction
--SELECT * FROM dbo.PLAN_DriverDocumentActionType
--SELECT * FROM dbo.IMP_Pracownicy
INSERT INTO dbo.PLAN_Driver
( IDNumber ,IDNumberHRSystem ,FirstName ,LastName ,Address ,ZipCode ,Place_ID ,Phone ,MobilePhone ,Email ,
DateOfBirth ,IDCardNumber ,IssuedBy ,PassportNumber ,PESEL ,Company_ID ,DriverGroup_ID ,DriverType_ID ,DriverMaster_ID ,User_ID ,ChangeDesc ,
CREATED ,MODIFIED ,Position ,UserPosition_ID ,CompanyGov_ID ,HRSystemStartDate ,HRSystemEndDate ,ColGUID ,Citizenship ,Description ,
ProfileDriverNo ,ProfileCreatedDate ,ProfileDeletedDate ,ProfileStatus ,SIParent_ID ,DriverTicketRegPrefSet_ID ,PIN ,
Deleted ,SuspendedDateFrom ,SuspendedDateTo ,BirthPlace ,MailAddress ,MailZipCode ,MailPlace_ID ,CompanyAddress_ID ,BankAccountNo ,
NIP ,DataUse ,IDCardValidDate)
SELECT  NR_SLUZB, ISNULL(NRSPLAC,NR_SLUZB), ISNULL(IMIE,'brak'), ISNULL(NAZWISKO,'brak'), ADRES, KODP, pl.ID, TEL, TELKOM, NULL,
DATA_UR, NR_DOWOD, NULL, NR_PASZ, p.PESEL, 1, NULL, CASE p.STATUS WHEN 7 THEN 1 WHEN 8 THEN 6 END, NULL, 1, 'import z PP',
GETDATE(), GETDATE(), '', NULL, NULL, NULL,NULL, NEWID(), NULL, NULL,
NULL, NULL, NULL, NULL, NULL, CASE p.STATUS WHEN 7 THEN 1 ELSE NULL END, CASE p.STATUS WHEN 7 THEN 1111 ELSE NULL END,
0,NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 
NULL, NULL, NULL
FROM dbo.IMP_Pracownicy p
LEFT JOIN dbo.TT_Place pl
ON p.SYM=pl.SYMBOL
AND p.MIEJSC=pl.NAME
LEFT JOIN dbo.PLAN_Driver d
ON p.NR_SLUZB = d.IDNumber
WHERE 
d.ID IS NULL


INSERT INTO dbo.PLAN_DriverAction
( Driver_ID ,DriverActionType_ID ,Date ,ValidTo ,Rate ,EmploymentType_ID ,MinBreakTime ,MaxBreakTime ,
Name ,Description ,CREATED ,MODIFIED ,ContractType_ID ,Company_ID ,ColGUID)
SELECT d.ID, 1, DATA_PRZYJ, DATA_ZWOL, CASE WHEN ISNUMERIC(ETAT)=1 THEN ETAT ELSE NULL END,1,NULL, NULL,
'','',GETDATE(),GETDATE(),8,1,NEWID()
FROM dbo.IMP_Pracownicy p
INNER JOIN dbo.PLAN_Driver d
ON p.NR_SLUZB = d.IDNumber
LEFT JOIN dbo.PLAN_DriverAction a
ON d.ID = a.Driver_ID
AND a.DriverActionType_ID=1
WHERE 
a.ID IS NULL




--nie s¹ importowane:
--przydzielona bileterka: NRINW, LOGO, DATPRZYBIL
--miejsce noclegowe (brak struktur w PLAN_Driver)
--urlop: DriverActionType_ID=3 (chyba bez sensu)
--pierwsze i ostatnie rozliczenie pracownika (nie bêdzie)
--prawo jazdy: DriverDocumentActionType_ID=12 (NUMERPRJ, DATAWYDPRJ, DATAWAZPRJ)
--badanie lekarskie: DriverDocumentActionType_ID=7
--badanie psychologiczne: DriverDocumentActionType_ID=8
--karta tachografu: DriverDocumentActionType_ID=11
--kurs w zakresie przewozu osób: DriverDocumentActionType_ID=9
--kurs w zakresie przewozu rzeczy: DriverDocumentActionType_ID=10
--SELECT * FROM dbo.IMP_Pracownicy
--SELECT * FROM dbo.PLAN_DriverType
--SELECT * FROM dbo.PLAN_DriverActionType
--SELECT * FROM dbo.PLAN_DriverAction
--SELECT * FROM dbo.PLAN_DriverDocumentActionType














GO
DROP TABLE dbo.IMP_Pracownicy;
GO


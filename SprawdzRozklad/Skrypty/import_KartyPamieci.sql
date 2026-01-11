USE [%DB%]
GO

--tworzenie tabeli
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[IMP_KartyPamieci]') AND type in (N'U'))
DROP TABLE [dbo].[IMP_KartyPamieci]
GO

CREATE TABLE IMP_KartyPamieci
(
NRKP varchar(max),
DATAPROD varchar(max),
POJEMN varchar(max),
TYP varchar(max),
BILETERKA varchar(max),
TYPBIL varchar(max),
ODDZ varchar(max),
DATAREJ varchar(max),
GODZREJ varchar(max),
NRPRACREJ varchar(max),
PINKARTY varchar(max),
NRSLUZBP varchar(max),
DATAZMPIN varchar(max),
NRKIER varchar(max),
IMIE varchar(max),
NAZWISKO varchar(max),
STATUS varchar(max),
STKP1 varchar(max),
STKP2 varchar(max),
MAXLPAS varchar(max),
NRWSIECI varchar(max),
KASA varchar(max),
LDNIREJ varchar(max),
DATAZMST varchar(max),
DATAWST varchar(max),
NROPST varchar(max),
DATAZAPST varchar(max),
GODZZAPST varchar(max),
BLOKADA varchar(max),
DATAWBLOK varchar(max),
NROPWB varchar(max),
DATAZBLOK varchar(max),
NROPZB varchar(max),
LZZBAKP varchar(max),
NAZWAZBA varchar(max),
DATAZZBA varchar(max),
GODZZZBA varchar(max),
NROPZZBA varchar(max),
NRPZAD varchar(max),
LZAD varchar(max),
PDZAD varchar(max),
ODZAD varchar(max),
NRKARDR varchar(max),
DATAZZAD varchar(max),
GODZZZAD varchar(max),
NROPZZAD varchar(max),
LZZADKP varchar(max),
DATAZIN varchar(max),
GODZZIN varchar(max),
NROPZIN varchar(max),
LZINKP varchar(max),
NAZPRZ varchar(max),
NRPRZ varchar(max),
DATAREJP varchar(max),
GODZREJP varchar(max),
DATAPOCZP varchar(max),
DATAZAKP varchar(max),
NRPRF varchar(max),
LOGOPRF varchar(max),
NRPZADKK varchar(max),
NRPKARDR varchar(max),
NRPRACPZ varchar(max),
IMIEPRPZ varchar(max),
NAZWPRPZ varchar(max),
STATPRPZ varchar(max),
NAZORZ varchar(max),
NRORZ varchar(max),
DATAREJO varchar(max),
GODZREJO varchar(max),
DATAPOCZO varchar(max),
DATAZAKO varchar(max),
NRORF varchar(max),
LOGOORF varchar(max),
NROZAD varchar(max),
NROKARDR varchar(max),
NRPRACOZ varchar(max),
IMIEPROZ varchar(max),
NAZWPROZ varchar(max),
STATPROZ varchar(max),
STKP3 varchar(max),
LDNIBLOK varchar(max),
STKP4 varchar(max),
INFOK1 varchar(max),
INFOK2 varchar(max),
MINDWOD varchar(max),
MAXPRZ varchar(max),
DATAWYCOF varchar(max),
DATALIKW varchar(max),
ST205_1 varchar(max),
DATAOP varchar(max),
GODZOP varchar(max),
NR_SLUZBOP varchar(max)
);


--START

--KONIEC  
GO
--usuniêcie starych - bileterka EMAR-05
DELETE FROM dbo.IMP_KartyPamieci WHERE (BILETERKA='EMAR-05') OR (BILETERKA is null)

update IMP_KartyPamieci
  set ST205_1=255
  where ST205_1 is NULL

--TUTAJ: START POPRAWEK DANYCH miejsce na kod do uporzadkowania danych: usuniêcie niepotrzebnych, poprawki, itp.

--SELECT * FROM dbo.IMP_KartyPamieci

--podwójne numery KP: nie mo¿e byæ! numer KP u¿ywam jako klucz przy imporcie
--SELECT NRKP, COUNT(1) FROM dbo.IMP_KartyPamieci GROUP BY NRKP HAVING COUNT(1)>1
--SELECT * FROM dbo.IMP_KartyPamieci WHERE NRKP='572FL'
--DELETE FROM dbo.IMP_KartyPamieci WHERE NRKP='???' AND NAZWISKO='???'

--usuniêcie starych - do zdefiniowania
--DELETE FROM dbo.IMP_KartyPamieci WHERE ???

--KONIEC!!!! POPRAWEK DANYCH


--import danych - MAT_TicketRegisterCard
--SELECT * FROM dbo.IMP_KartyPamieci
--SELECT * FROM dbo.MAT_TicketRegisterCard

INSERT INTO dbo.MAT_TicketRegisterCard
( CardSN ,ProductionDate ,Capacity ,TicketRegisterType ,DataType ,Company_ID ,
RegistrationDate ,Driver_ID ,Blocked ,BlockedDate ,UnblockedDate ,ScrapDate ,
Description ,CREATED ,MODIFIED)
SELECT NRKP, ISNULL(convert(DATETIME,DATAPROD,104),'2017-01-01'),CASE POJEMN WHEN 15 THEN 2048 ELSE 0 END, BILETERKA, TYPBIL, NULL,
ISNULL(convert(DATETIME,DATAREJ,104),'2017-01-01'), d.ID, ISNULL(BLOKADA,0), convert(DATETIME,DATAWBLOK,104), convert(DATETIME,DATAZBLOK,104), NULL, 
'import z PP/KK', GETDATE(), GETDATE()
FROM dbo.IMP_KartyPamieci k
LEFT JOIN dbo.PLAN_Driver d
ON k.NRKIER = d.IDNumber and d.DriverType_ID=1
LEFT JOIN dbo.MAT_TicketRegisterCard c
ON c.CardSN=k.NRKP
WHERE c.ID IS NULL

--SELECT * FROM dbo.MAT_DriverTicketRegisterCardHistory
INSERT INTO dbo.MAT_DriverTicketRegisterCardHistory
( Driver_ID ,TicketRegisterCard_ID ,ValidFrom ,ValidTo ,CREATED ,MODIFIED)
SELECT c.Driver_ID,c.ID,c.RegistrationDate,NULL, GETDATE(), GETDATE()
FROM dbo.IMP_KartyPamieci k
INNER JOIN dbo.MAT_TicketRegisterCard c
ON c.CardSN=k.NRKP
LEFT JOIN dbo.MAT_DriverTicketRegisterCardHistory h
ON h.TicketRegisterCard_ID = c.ID
WHERE c.Driver_ID IS NOT NULL
and h.ID IS NULL

--zmiana PIN-ów kierowców
UPDATE dbo.PLAN_Driver
SET PIN= PINKARTY
FROM dbo.IMP_KartyPamieci k
INNER JOIN dbo.MAT_TicketRegisterCard c
ON c.CardSN=k.NRKP
INNER JOIN dbo.PLAN_Driver d
ON c.Driver_ID = d.ID

--SELECT * FROM dbo.IMP_KartyPamieci
--SELECT * FROM dbo.PLAN_DriverTicketRegPrefSet
INSERT INTO dbo.PLAN_DriverTicketRegPrefSet
( Name ,ReadOnly ,GroupSet ,Description 
,[10NoClockChange],[11NoFeesAllowed] ,[12TaskRepLimit] ,[13NoTaskRepAllowed] ,[14NoControlTicketAllowed] , --1
[15NoPrevBusStop] ,[16NoGroupTicket] ,[17BusStopNumberFromLine] --1
,[20AllowRideWithoutCalendar] ,[21AllowControlNumCheck] ,[22CardFullReminder] , [23DisallowPINChange] ,[24DisallowTRChange] , --2
[25PassengerDataEMK],[26NoRideReportPrint] ,[27NoPeriodReportPrint] --2
,[30NoBusNumberOnTicket] ,[31NoForeignCurrency] , [32NoForcedMonthReport] ,[33NoFiscalReportReminder] ,[34AlwaysNewBusNumber] , --3
[35BusNumberOnDemand] ,[36BusNumberForbidden] ,[37ZoneTicketSaleAsteriks] , --3
[40PrintLineData] ,[41ForeignCurrencyReceipt] ,[42NoNextPeriodSales] ,[43ControlDeviceWithoutPrint] ,[44AllowRideWithoutTask] --4
,[45AllowReductionNoNumber] ,[46AllowIncomeDisplay] ,[47NoQuickSaleMode], [EMAR205PaymentMethod] --4
,[AllowZeroFiscalReports],[IncludeSkippedStops],         --O205_1 bity 4,5    
[CardPaymentWithoutPaymentTerminal],[UsePaymentTerminal] --O205_1 bity 6,7
,DateStatusChanged ,MaxPassangerOnTicket ,DaysBeforePrompt ,DaysBeforeBlock ,CREATED ,MODIFIED ,DefaultSet ,InfoLine1 ,InfoLine2)
SELECT 'Import PP/KK, kier,nrkp: ' + ISNULL(NRKIER,'???')+','+ISNULL(k.NRKP,'???'),0,1,NULL,
--stkp1,dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),
-- STKP1 OIK.354
case when RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),1)=0 then 1 else 0 end, --bit 0
--RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),1),
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),2),1)=0 then 1 else 0 end,  --bit 1
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),2),1),  --bit 1
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),3),1)=0 then 1 else 0 end,
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),3),1),
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),4),1)=0 then 1 else 0 end,
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),4),1),
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),5),1)=0 then 1 else 0 end, 
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),5),1), 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),6),1)=0 then 1 else 0 end, 
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),6),1), 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),7),1)=0 then 1 else 0 end, 
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),7),1), 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),8),1)=0 then 1 else 0 end,
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP1 AS INT)),8),1),
--stkp2,dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),
-- STKP2 OIK.355
RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),1),  --bit 0
--case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),2),1)=0 then 1 else 0 end,  --bit 1
LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),2),1),  --bit 1
--case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),3),1)=0 then 1 else 0 end,   --bit 2
LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),3),1),   --bit 2
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),4),1)=0 then 1 else 0 end,
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),4),1),
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),5),1)=0 then 1 else 0 end, 
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),5),1), 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),6),1)=0 then 1 else 0 end, 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),7),1)=0 then 1 else 0 end, 
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),7),1), 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),8),1)=0 then 1 else 0 end,
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP2 AS INT)),8),1),
--stkp3,dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),
-- STKP3 OIK.391
case when RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),1)=0 then 1 else 0 end, --bit 0
--RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),1),   --bit 0
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),2),1)=0 then 1 else 0 end, --bit 1
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),2),1),
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),3),1)=0 then 1 else 0 end,  --bit 2
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),3),1),
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),4),1)=0 then 1 else 0 end,  -- bit 3
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),4),1),
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),5),1)=0 then 1 else 0 end,   --bit 4
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),5),1), 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),6),1)=0 then 1 else 0 end,  --bit 5
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),6),1), 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),7),1)=0 then 1 else 0 end,   --bit 6
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),7),1), 
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),8),1)=0 then 1 else 0 end,  --bit 7
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP3 AS INT)),8),1),
--stkp4,dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),
-- STKP4 OIK.421
case when RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),1)=0 then 1 else 0 end, --bit 0
--RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),1),
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),2),1)=0 then 1 else 0 end,  --bit 1
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),3),1)=0 then 1 else 0 end, --bit 2
--LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),3),1),  --bit 2
--case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),4),1)=0 then 1 else 0 end, --bit 3
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),4),1)=0 then 1 else 0 end,  --bit 3
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),5),1)=0 then 1 else 0 end,  --bit 4
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),6),1)=0 then 1 else 0 end,  --bit 5
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),7),1)=0 then 1 else 0 end,  --bit 6
--Uwaga: w bileterce bit 7 interpretowany jest jako: 0-dozwolone/1-zabronione przyjmowanie zap³aty bezgotówkowo
--case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),8),1)=0 then 1 else 0 end,
0,      --bit 7 (pole [47NoQuickSaleMode] nie jest obs³ugiwane w Informica 2.0) - w KK interpretacja tego pola by³a zmieniona na
-- 0-dozwolone przyjmowanie zap³aty bezgotówkowo
-- 1-zabronione przyjmowanie zap³aty bezgotówkowo
-- dotyczy to tylko EMAR-205 - bit 7 STKP4 ustawiany jest przez pole [EMAR205PaymentMethod]
--case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),8),1)=0 then 1 else 0 end, --[EMAR205PaymentMethod]
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(STKP4 AS INT)),8),1)=0 then 1 else 0 end,
--ST205_1 --> O205_1 OIK.492
--[AllowZeroFiscalReports],[IncludeSkippedStops],         --O205_1 bity 4,5    
--[CardPaymentWithoutPaymentTerminal],[UsePaymentTerminal] --O205_1 bity 6,7
LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(ST205_1 AS INT)),5),1),  --bit 4
LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(ST205_1 AS INT)),6),1),  --bit 5 
LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(ST205_1 AS INT)),7),1),  --bit 6
case when LEFT(RIGHT(dbo.UTIL_BinaryStringFromInt(CAST(ST205_1 AS INT)),8),1)=0 then 1 else 0 end,  --bit 7
convert(DATETIME,DATAZMST,104),MAXLPAS,LDNIREJ,LDNIBLOK, GETDATE(), GETDATE(),0, INFOK1,INFOK2
FROM dbo.IMP_KartyPamieci k
INNER JOIN dbo.MAT_TicketRegisterCard c
ON c.CardSN=k.NRKP
INNER JOIN dbo.PLAN_Driver d
ON c.Driver_ID = d.ID
LEFT JOIN dbo.PLAN_DriverTicketRegPrefSet s
ON s.Name = 'Import PP/KK, kier,nrkp: ' + ISNULL(NRKIER,'???')+','+ISNULL(k.NRKP,'???')
WHERE 
s.ID IS NULL

UPDATE dbo.PLAN_Driver
SET DriverTicketRegPrefSet_ID = s.ID
FROM dbo.IMP_KartyPamieci k
INNER JOIN dbo.MAT_TicketRegisterCard c
ON c.CardSN=k.NRKP
INNER JOIN dbo.PLAN_Driver d
ON c.Driver_ID = d.ID
INNER JOIN dbo.PLAN_DriverTicketRegPrefSet s
ON s.Name = 'Import PP/KK, kier,nrkp: ' + ISNULL(NRKIER,'???')+','+ISNULL(k.NRKP,'???')


UPDATE PLAN_DriverTicketRegPrefSet
SET GroupSet = 0
WHERE ID IN (
    -- Wybiera ID z PLAN_DriverTicketRegPrefSet, dla których
    -- DriverTicketRegPrefSet_ID wystêpuje tylko raz w PLAN_Driver
    SELECT T1.ID
    FROM PLAN_DriverTicketRegPrefSet T1
    JOIN PLAN_Driver T2
        ON T1.ID = T2.DriverTicketRegPrefSet_ID
    GROUP BY T1.ID
    HAVING COUNT(T2.DriverTicketRegPrefSet_ID) = 1
);

GO
DROP TABLE dbo.IMP_KartyPamieci;
GO

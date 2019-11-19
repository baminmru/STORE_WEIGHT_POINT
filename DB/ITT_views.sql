SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


ALTER   view V_viewITTIN_ITTIN_PALET as 
select   ITTIN_PALETID, ITTIN_QLINEID
, 
ITTIN_QLINE.exp_date 
ITTIN_QLINE_exp_date 
, 
 ITTIN_PALET.IsBrak  
ITTIN_PALET_IsBrak_VAL, 
 case ITTIN_PALET.IsBrak 
when -1 then 'Да'
when 0 then 'Нет'
 end 
ITTIN_PALET_IsBrak 
, 
ITTIN_DEF.StampNumber 
ITTIN_DEF_StampNumber 
, 
 dbo.GetBriefFromXML(ITTIN_QLINE.good_id) 
ITTIN_QLINE_good_id 
, 
ITTIN_PALET.BufferZonePlace 
ITTIN_PALET_BufferZonePlace 
, 
 ITTIN_PALET.made_country  
ITTIN_PALET_made_country_ID, 
 dbo.ITTD_COUNTRY_BRIEF_F(ITTIN_PALET.made_country,null) 
ITTIN_PALET_made_country 
, 
ITTIN_DEF.TranspNumber 
ITTIN_DEF_TranspNumber 
, 
ITTIN_QLINE.FullPackageWeight 
ITTIN_QLINE_FullPackageWeight 
, 
ITTIN_PALET.CaliberQuantity 
ITTIN_PALET_CaliberQuantity 
, 
 dbo.GetBriefFromXML(ITTIN_QLINE.QRY_NUM) 
ITTIN_QLINE_QRY_NUM 
, 
ITTIN_PALET.FullPackageWeight 
ITTIN_PALET_FullPackageWeight 
, 
ITTIN_DEF.TTN 
ITTIN_DEF_TTN 
, 
ITTIN_PALET.PackageWeight 
ITTIN_PALET_PackageWeight 
, 
 dbo.GetBriefFromXML(ITTIN_QLINE.LineAtQuery) 
ITTIN_QLINE_LineAtQuery 
, 
ITTIN_PALET.palet_id 
ITTIN_PALET_palet_id 
, 
ITTIN_QLINE.articul 
ITTIN_QLINE_articul 
, 
ITTIN_PALET.Made_date 
ITTIN_PALET_Made_date 
, 
 ITTIN_QLINE.PartRef  
ITTIN_QLINE_PartRef_ID, 
 dbo.ITTD_PART_BRIEF_F(ITTIN_QLINE.PartRef,null) 
ITTIN_QLINE_PartRef 
, 
 ITTIN_QLINE.KILL_NUMBER  
ITTIN_QLINE_KILL_NUMBER_ID, 
 dbo.ITTD_KILLPLACE_BRIEF_F(ITTIN_QLINE.KILL_NUMBER,null) 
ITTIN_QLINE_KILL_NUMBER 
, 
 ITTIN_QLINE.Navalom  
ITTIN_QLINE_Navalom_VAL, 
 case ITTIN_QLINE.Navalom 
when -1 then 'Да'
when 0 then 'Нет'
 end 
ITTIN_QLINE_Navalom 
, 
ITTIN_PALET.Stock_ID 
ITTIN_PALET_Stock_ID 
, 
ITTIN_PALET.KorobNetto 
ITTIN_PALET_KorobNetto 
, 
ITTIN_PALET.exp_date 
ITTIN_PALET_exp_date 
, 
ITTIN_QLINE.edizm 
ITTIN_QLINE_edizm 
, 
ITTIN_QLINE.CaliberWeight 
ITTIN_QLINE_CaliberWeight 
, 
 dbo.GetBriefFromXML(ITTIN_DEF.QryCode) 
ITTIN_DEF_QryCode 
, 
ITTIN_DEF.ProcessDate 
ITTIN_DEF_ProcessDate 
, 
ITTIN_DEF.StampStatus 
ITTIN_DEF_StampStatus 
, 
 ITTIN_PALET.PartRef  
ITTIN_PALET_PartRef_ID, 
 dbo.ITTD_PART_BRIEF_F(ITTIN_PALET.PartRef,null) 
ITTIN_PALET_PartRef 
, 
ITTIN_QLINE.VidOtruba 
ITTIN_QLINE_VidOtruba 
, 
ITTIN_DEF.Container 
ITTIN_DEF_Container 
, 
ITTIN_DEF.Track_time_in 
ITTIN_DEF_Track_time_in 
, 
ITTIN_QLINE.CurValue 
ITTIN_QLINE_CurValue 
, 
ITTIN_QLINE.sequence 
ITTIN_QLINE_sequence 
, 
 ITTIN_QLINE.IsCalibrated  
ITTIN_QLINE_IsCalibrated_VAL, 
 case ITTIN_QLINE.IsCalibrated 
when -1 then 'Да'
when 0 then 'Нет'
 end 
ITTIN_QLINE_IsCalibrated 
, 
ITTIN_QLINE.KorobBrutto 
ITTIN_QLINE_KorobBrutto 
, 
ITTIN_DEF.TTNDate 
ITTIN_DEF_TTNDate 
, 
ITTIN_DEF.track_time_out 
ITTIN_DEF_track_time_out 
, 
ITTIN_PALET.KorobBrutto 
ITTIN_PALET_KorobBrutto 
, 
 ITTIN_PALET.TheNumber  
ITTIN_PALET_TheNumber_ID, 
 dbo.ITTPL_DEF_BRIEF_F(ITTIN_PALET.TheNumber,null) 
ITTIN_PALET_TheNumber 
, 
 dbo.GetBriefFromXML(ITTIN_DEF.TheClient) 
ITTIN_DEF_TheClient 
, 
ITTIN_PALET.VidOtruba 
ITTIN_PALET_VidOtruba 
, 
 ITTIN_QLINE.made_country  
ITTIN_QLINE_made_country_ID, 
 dbo.ITTD_COUNTRY_BRIEF_F(ITTIN_QLINE.made_country,null) 
ITTIN_QLINE_made_country 
, 
 ITTIN_PALET.KILL_NUMBER  
ITTIN_PALET_KILL_NUMBER_ID, 
 dbo.ITTD_KILLPLACE_BRIEF_F(ITTIN_PALET.KILL_NUMBER,null) 
ITTIN_PALET_KILL_NUMBER 
, 
ITTIN_QLINE.PackageWeight 
ITTIN_QLINE_PackageWeight 
, 
ITTIN_QLINE.Made_date 
ITTIN_QLINE_Made_date 
, 
 ITTIN_PALET.Factory  
ITTIN_PALET_Factory_ID, 
 dbo.ITTD_FACTORY_BRIEF_F(ITTIN_PALET.Factory,null) 
ITTIN_PALET_Factory 
, 
ITTIN_PALET.sequence 
ITTIN_PALET_sequence 
, 
ITTIN_PALET.PalWeight 
ITTIN_PALET_PalWeight 
, 
ITTIN_PALET.GoodWithPaletWeight 
ITTIN_PALET_GoodWithPaletWeight 
, 
ITTIN_DEF.Supplier 
ITTIN_DEF_Supplier 
, 
ITTIN_QLINE.KorobNetto 
ITTIN_QLINE_KorobNetto 
, 
ITTIN_DEF.temp_in_track 
ITTIN_DEF_temp_in_track 
, 
 ITTIN_QLINE.Factory  
ITTIN_QLINE_Factory_ID, 
 dbo.ITTD_FACTORY_BRIEF_F(ITTIN_QLINE.Factory,null) 
ITTIN_QLINE_Factory 
,
 ITTIN_PALET.IsCalibrated  
ITTIN_PALET_IsCalibrated_VAL, 
 case ITTIN_PALET.IsCalibrated 
when -1 then 'Да'
when 0 then 'Нет'
 end 
ITTIN_PALET_IsCalibrated 

, ITTIN_QLINE.InstanceID InstanceID 
, ITTIN_PALET.ITTIN_PALETID ID 
, 'ITTIN_PALET' VIEWBASE 
, XXXMYSTATUSXXX.Name StatusName 
, XXXMYSTATUSXXX.objstatusid INTSANCEStatusID
,ITTPL_DEF.WEIGHT poddonweight
,ITTPL_DEF.PrivatePalet
,dbo.ITTD_PLTYPE_BRIEF_F(ITTPL_DEF.PLTYPE,null) ITTPL_DEF_PLTYPE

 from ITTIN_PALET
 join ITTIN_QLINE on ITTIN_QLINE.ITTIN_QLINEID=ITTIN_PALET.ParentStructRowID 
 join INSTANCE on ITTIN_QLINE.INSTANCEID=INSTANCE.INSTANCEID
 left join objstatus XXXMYSTATUSXXX on instance.status=XXXMYSTATUSXXX.objstatusid
 left join ITTIN_DEF ON ITTIN_DEF.InstanceID=ITTIN_QLINE.InstanceID
 left join ITTPL_DEF on ITTPL_DEF.ITTPL_DEFID =ITTIN_PALET.THENUMBER


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



ALTER   view V_viewITTOUT_ITTOUT_PALET as 
select   ITTOUT_PALETID, ITTOUT_LINESID
, 
ITTOUT_DEF.TTNDate 
ITTOUT_DEF_TTNDate 
, 
ITTOUT_DEF.temp_in_track 
ITTOUT_DEF_temp_in_track 
, 
ITTOUT_DEF.track_time_out 
ITTOUT_DEF_track_time_out 
, 
ITTOUT_DEF.TranspNumber 
ITTOUT_DEF_TranspNumber 
, 
 dbo.GetBriefFromXML(ITTOUT_DEF.ShipOrder) 
ITTOUT_DEF_ShipOrder 
, 
ITTOUT_DEF.StampNumber 
ITTOUT_DEF_StampNumber 
, 
ITTOUT_DEF.StampStatus 
ITTOUT_DEF_StampStatus 
, 
ITTOUT_DEF.Supplier 
ITTOUT_DEF_Supplier 
, 
 dbo.GetBriefFromXML(ITTOUT_DEF.TheClient) 
ITTOUT_DEF_TheClient 
, 
ITTOUT_DEF.TTN 
ITTOUT_DEF_TTN 
, 
ITTOUT_DEF.ProcessDate 
ITTOUT_DEF_ProcessDate 
, 
ITTOUT_DEF.Container 
ITTOUT_DEF_Container 
, 
ITTOUT_DEF.Track_time_in 
ITTOUT_DEF_Track_time_in 
, 
ITTOUT_LINES.sequence 
ITTOUT_LINES_sequence 
, 
 ITTOUT_LINES.Navalom  
ITTOUT_LINES_Navalom_VAL, 
 case ITTOUT_LINES.Navalom 
when -1 then 'Да'
when 0 then 'Нет'
 end 
ITTOUT_LINES_Navalom 
, 
 ITTOUT_PALET.IsBrak  
ITTOUT_PALET_IsBrak_VAL, 
 case ITTOUT_PALET.IsBrak
when -1 then 'Да'
when 0 then 'Нет'
 end 
ITTOUT_PALET_IsBrak 
, 
 ITTOUT_LINES.Factory  
ITTOUT_LINES_Factory_ID, 
 dbo.ITTD_FACTORY_BRIEF_F(ITTOUT_LINES.Factory,null) 
ITTOUT_LINES_Factory 
, 
ITTOUT_LINES.Quanity 
ITTOUT_LINES_Quanity 
, 
ITTOUT_LINES.exp_date 
ITTOUT_LINES_exp_date 
, 
 ITTOUT_LINES.PartRef  
ITTOUT_LINES_PartRef_ID, 
 dbo.ITTD_PART_BRIEF_F(ITTOUT_LINES.PartRef,null) 
ITTOUT_LINES_PartRef 
, 
ITTOUT_LINES.FullPackageWeight 
ITTOUT_LINES_FullPackageWeight 
, 
ITTOUT_LINES.Made_date 
ITTOUT_LINES_Made_date 
, 
 ITTOUT_LINES.KILL_NUMBER  
ITTOUT_LINES_KILL_NUMBER_ID, 
 dbo.ITTD_KILLPLACE_BRIEF_F(ITTOUT_LINES.KILL_NUMBER,null) 
ITTOUT_LINES_KILL_NUMBER 
, 
ITTOUT_LINES.articul 
ITTOUT_LINES_articul 
, 
ITTOUT_LINES.PackageWeight 
ITTOUT_LINES_PackageWeight 
, 
 ITTOUT_LINES.made_country  
ITTOUT_LINES_made_country_ID, 
 dbo.ITTD_COUNTRY_BRIEF_F(ITTOUT_LINES.made_country,null) 
ITTOUT_LINES_made_country 
, 
 dbo.GetBriefFromXML(ITTOUT_LINES.LineAtQuery) 
ITTOUT_LINES_LineAtQuery 
, 
 dbo.GetBriefFromXML(ITTOUT_LINES.good_ID) 
ITTOUT_LINES_good_ID 
, 
 dbo.GetBriefFromXML(ITTOUT_LINES.QRY_NUM) 
ITTOUT_LINES_QRY_NUM 
, 
ITTOUT_LINES.VidOtruba 
ITTOUT_LINES_VidOtruba 
, 
ITTOUT_LINES.edizm 
ITTOUT_LINES_edizm 
, 
ITTOUT_LINES.NumInBufZone 
ITTOUT_LINES_NumInBufZone 
, 
ITTOUT_LINES.CurValue 
ITTOUT_LINES_CurValue 
, 
ITTOUT_PALET.StoreCell 
ITTOUT_PALET_StoreCell 
, 
ITTOUT_PALET.PackageWeight 
ITTOUT_PALET_PackageWeight 
, 
ITTOUT_PALET.FullPackageWeight 
ITTOUT_PALET_FullPackageWeight 
, 
ITTOUT_PALET.VidOtruba 
ITTOUT_PALET_VidOtruba 
, 
ITTOUT_PALET.BufferCell 
ITTOUT_PALET_BufferCell 
, 
ITTOUT_PALET.exp_date 
ITTOUT_PALET_exp_date 
, 
ITTOUT_PALET.CaliberQuantity 
ITTOUT_PALET_CaliberQuantity 
, 
ITTOUT_PALET.Made_date 
ITTOUT_PALET_Made_date 
, 
 ITTOUT_PALET.IsEmpty  
ITTOUT_PALET_IsEmpty_VAL, 
 case ITTOUT_PALET.IsEmpty 
when -1 then 'Да'
when 0 then 'Нет'
 end 
ITTOUT_PALET_IsEmpty 
,
 ITTOUT_PALET.IsCalibrated  
ITTOUT_PALET_IsCalibrated_VAL, 
 case ITTOUT_PALET.IsCalibrated 
when -1 then 'Да'
when 0 then 'Нет'
 end 
ITTOUT_PALET_IsCalibrated 
,
ITTOUT_PALET.ReorgPackageFullWeight 
ITTOUT_PALET_ReorgPackageFullWeight 
, 
 ITTOUT_PALET.PartRef  
ITTOUT_PALET_PartRef_ID, 
 dbo.ITTD_PART_BRIEF_F(ITTOUT_PALET.PartRef,null) 
ITTOUT_PALET_PartRef 
, 
 ITTOUT_PALET.KILL_NUMBER  
ITTOUT_PALET_KILL_NUMBER_ID, 
 dbo.ITTD_KILLPLACE_BRIEF_F(ITTOUT_PALET.KILL_NUMBER,null) 
ITTOUT_PALET_KILL_NUMBER 
, 
 ITTOUT_PALET.Factory  
ITTOUT_PALET_Factory_ID, 
 dbo.ITTD_FACTORY_BRIEF_F(ITTOUT_PALET.Factory,null) 
ITTOUT_PALET_Factory 
, 
ITTOUT_PALET.GoodWithPaletWeight 
ITTOUT_PALET_GoodWithPaletWeight 
, 
ITTOUT_PALET.sequence 
ITTOUT_PALET_sequence 
, 
 ITTOUT_PALET.made_country  
ITTOUT_PALET_made_country_ID, 
 dbo.ITTD_COUNTRY_BRIEF_F(ITTOUT_PALET.made_country,null) 
ITTOUT_PALET_made_country 
, 
ITTOUT_PALET.ReorgWeight 
ITTOUT_PALET_ReorgWeight 
, 
 ITTOUT_PALET.TheNumber  
ITTOUT_PALET_TheNumber_ID, 
 dbo.ITTPL_DEF_BRIEF_F(ITTOUT_PALET.TheNumber,null) 
ITTOUT_PALET_TheNumber 
, 
ITTOUT_PALET.ReorgCaliberQuantity 
ITTOUT_PALET_ReorgCaliberQuantity 
, ITTOUT_LINES.InstanceID InstanceID 
, ITTOUT_PALET.ITTOUT_PALETID ID 
, 'ITTOUT_PALET' VIEWBASE 
, XXXMYSTATUSXXX.Name StatusName 
, XXXMYSTATUSXXX.objstatusid INTSANCEStatusID
,ITTPL_DEF.WEIGHT poddonweight
,ITTPL_DEF.PrivatePalet

,ITTPL_DEF.QryInNumber
ITTPL_DEF_QryInNumber
,
ITTPL_DEF.QryInDate
ITTPL_DEF_QryInDate
,dbo.ITTD_PLTYPE_BRIEF_F(ITTPL_DEF.PLTYPE,null) ITTPL_DEF_PLTYPE


 from ITTOUT_PALET
 join ITTOUT_LINES on ITTOUT_LINES.ITTOUT_LINESID=ITTOUT_PALET.ParentStructRowID 
 join INSTANCE on ITTOUT_LINES.INSTANCEID=INSTANCE.INSTANCEID
 left join objstatus XXXMYSTATUSXXX on instance.status=XXXMYSTATUSXXX.objstatusid
 left join ITTOUT_DEF ON ITTOUT_DEF.InstanceID=ITTOUT_LINES.InstanceID
 left join ITTPL_DEF on ITTPL_DEF.ITTPL_DEFID =ITTOUT_PALET.THENUMBER


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


ALTER  view V_viewITTOUT_ITTOUT_SRV as 
select   ITTOUT_SRVID
, 
ITTOUT_DEF.ProcessDate 
ITTOUT_DEF_ProcessDate 
, 
ITTOUT_DEF.TTNDate 
ITTOUT_DEF_TTNDate 
, 
ITTOUT_DEF.TranspNumber 
ITTOUT_DEF_TranspNumber 
, 
 dbo.GetBriefFromXML(ITTOUT_DEF.ShipOrder) 
ITTOUT_DEF_ShipOrder 
, 
dbo.GetIDFromXML(ITTOUT_DEF.ShipOrder) 
ITTOUT_DEF_ShipOrder_ID 
, 
 dbo.GetBriefFromXML(ITTOUT_DEF.TheClient) 
ITTOUT_DEF_TheClient 
, 
 dbo.GetIDFromXML(ITTOUT_DEF.TheClient) 
ITTOUT_DEF_TheClient_ID
, 
ITTOUT_DEF.StampStatus 
ITTOUT_DEF_StampStatus 
, 
ITTOUT_DEF.TTN 
ITTOUT_DEF_TTN 
, 
ITTOUT_DEF.Supplier 
ITTOUT_DEF_Supplier 
, 
ITTOUT_DEF.StampNumber 
ITTOUT_DEF_StampNumber 
, 
ITTOUT_DEF.temp_in_track 
ITTOUT_DEF_temp_in_track 
, 
 ITTOUT_SRV.SRV  
ITTOUT_SRV_SRV_ID, 
 dbo.ITTD_SRV_BRIEF_F(ITTOUT_SRV.SRV,null) 
ITTOUT_SRV_SRV 
, 
ITTOUT_DEF.Container 
ITTOUT_DEF_Container 
, 
ITTOUT_DEF.Track_time_in 
ITTOUT_DEF_Track_time_in 
, 
ITTOUT_SRV.Quantity 
ITTOUT_SRV_Quantity 
, 
ITTOUT_DEF.track_time_out 
ITTOUT_DEF_track_time_out 
, ITTOUT_SRV.InstanceID InstanceID 
, ITTOUT_SRV.ITTOUT_SRVID ID 
, 'ITTOUT_SRV' VIEWBASE 
, XXXMYSTATUSXXX.Name StatusName 
, XXXMYSTATUSXXX.objstatusid INTSANCEStatusID

 from ITTOUT_SRV
 join INSTANCE on ITTOUT_SRV.INSTANCEID=INSTANCE.INSTANCEID
 left join objstatus XXXMYSTATUSXXX on instance.status=XXXMYSTATUSXXX.objstatusid
 left join ITTOUT_DEF ON ITTOUT_DEF.InstanceID=ITTOUT_SRV.InstanceID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  SELECT  ON [dbo].[V_viewITTOUT_ITTOUT_SRV]  TO [public]
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


alter view V_viewITTIN_ITTIN_SRV as 
select   ITTIN_SRVID
, 
ITTIN_DEF.TranspNumber 
ITTIN_DEF_TranspNumber 
, 
ITTIN_DEF.TTNDate 
ITTIN_DEF_TTNDate 
, 
 dbo.GetBriefFromXML(ITTIN_DEF.TheClient) 
ITTIN_DEF_TheClient 
, 
 dbo.GetIDFromXML(ITTIN_DEF.TheClient) 
ITTIN_DEF_TheClient_ID 
, 
 dbo.GetBriefFromXML(ITTIN_DEF.QryCode) 
ITTIN_DEF_QryCode 
, 
 dbo.GetIDFromXML(ITTIN_DEF.QryCode) 
ITTIN_DEF_QryCode_ID 
, 
ITTIN_DEF.Track_time_in 
ITTIN_DEF_Track_time_in 
, 
ITTIN_DEF.TTN 
ITTIN_DEF_TTN 
, 
ITTIN_DEF.ProcessDate 
ITTIN_DEF_ProcessDate 
, 
ITTIN_DEF.StampNumber 
ITTIN_DEF_StampNumber 
, 
 ITTIN_SRV.SRV  
ITTIN_SRV_SRV_ID, 
 dbo.ITTD_SRV_BRIEF_F(ITTIN_SRV.SRV,null) 
ITTIN_SRV_SRV 
, 
ITTIN_DEF.Container 
ITTIN_DEF_Container 
, 
ITTIN_DEF.temp_in_track 
ITTIN_DEF_temp_in_track 
, 
ITTIN_SRV.Quantity 
ITTIN_SRV_Quantity 
, 
ITTIN_DEF.StampStatus 
ITTIN_DEF_StampStatus 
, 
ITTIN_DEF.track_time_out 
ITTIN_DEF_track_time_out 
, 
ITTIN_DEF.Supplier 
ITTIN_DEF_Supplier 
, ITTIN_SRV.InstanceID InstanceID 
, ITTIN_SRV.ITTIN_SRVID ID 
, 'ITTIN_SRV' VIEWBASE 
, XXXMYSTATUSXXX.Name StatusName 
, XXXMYSTATUSXXX.objstatusid INTSANCEStatusID

 from ITTIN_SRV
 join INSTANCE on ITTIN_SRV.INSTANCEID=INSTANCE.INSTANCEID
 left join objstatus XXXMYSTATUSXXX on instance.status=XXXMYSTATUSXXX.objstatusid
 left join ITTIN_DEF ON ITTIN_DEF.InstanceID=ITTIN_SRV.InstanceID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  SELECT  ON [dbo].[V_viewITTIN_ITTIN_SRV]  TO [public]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_SERVICE]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view V_SERVICE
go

create   view V_SERVICE as
select 'Приемка' ZTYPE, ITTIN_DEF_TheClient as CLIENT,
ITTIN_DEF_TheClient_ID CLIENT_ID ,
ITTIN_SRV_SRV as SRV , 
ITTD_SRV.LinkCode SRV_ID,
ITTIN_SRV_Quantity as Quantity,ITTIN_DEF_QryCode as ZAKAZ,
ITTIN_DEF_QryCode_ID as  ZAKAZ_ID, ITTIN_DEF_ProcessDate as ProcessDate from v_viewITTIN_ITTIN_SRV
join ITTD_SRV on ITTD_SRVID=ITTIN_SRV_SRV_ID
union all
select 'Отгрузка',ITTOUT_DEF_TheClient,
 ITTOUT_DEF_TheClient_ID, ITTOUT_SRV_SRV,
ITTD_SRV.LinkCode ,
ITTOUT_SRV_Quantity,ITTOUT_DEF_ShipOrder,ITTOUT_DEF_ShipOrder_ID, ITTOUT_DEF_ProcessDate from v_viewITTOUT_ITTOUT_SRV
join ITTD_SRV on ITTD_SRVID=ITTOUT_SRV_SRV_ID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


update ittpl_def set pltype='C606C48F-E7AF-42A5-91E2-17A04B21262B' where pltype  is null
go




delete from ROLES_USER
go
delete from ROLES_MAP
go
delete from ROLES_DOC_STATE
go
delete from ROLES_DOC
go
delete from ROLES_DEF
go
delete from ROLES_REPORTS
go
delete from ROLES_ACT
go
delete from ROLES_WP
go
delete from instance where objtype='ROLES'
go
delete from FIELDTYPEMAP
go
delete from ENUMITEM
go
delete from FIELDTYPE
go
delete from PARAMETERS
go
delete from SCRIPT
go
delete from SHAREDMETHOD
go
delete from PARTPARAMMAP
go
delete from PARTMENU
go
delete from FIELDVALIDATOR
go
delete from FIELDPARAMMAP
go
delete from FIELDMENU
go
delete from FldExtenders
go
delete from FIELDEXPRESSION
go
delete from DINAMICFILTERSCRIPT
go
delete from FIELDSRCDEF
go
delete from FIELD
go
delete from ViewColumn
go
delete from PARTVIEW_LNK
go
delete from PARTVIEW
go
delete from ExtenderInterface
go
delete from VALIDATOR
go
delete from CONSTRAINTFIELD
go
delete from UNIQUECONSTRAINT
go
delete from PART
go
delete from INSTANCEVALIDATOR
go
delete from NEXTSTATE
go
delete from OBJSTATUS
go
delete from STRUCTRESTRICTION
go
delete from FIELDRESTRICTION
go
delete from METHODRESTRICTION
go
delete from OBJECTMODE
go
delete from TYPEMENU
go
delete from OBJECTTYPE
go
delete from ParentPackage
go
delete from MTZAPP
go
delete from GENMANUALCODE
go
delete from GENCONTROLS
go
delete from GENREFERENCE
go
delete from GENERATOR_TARGET
go
delete from GENPACKAGE
go
delete from LocalizeInfo
go
delete from instance where objtype='MTZMetaModel'
go
delete from WorkPlace
go
delete from EPFilterLink
go
delete from EntryPoints
go
delete from ARMTypes
go
delete from ARMJRNLADD
go
delete from ARMJRNLREP
go
delete from ARMJRNLRUN
go
delete from ARMJournal
go
delete from instance where objtype='MTZwp'
go
delete from JColumnSource
go
delete from JournalColumn
go
delete from Journal
go
delete from JournalSrc
go
delete from instance where objtype='MTZJrnl'
go
delete from FileterField
go
delete from FilterFieldGroup
go
delete from Filters
go
delete from instance where objtype='MTZFltr'
go

truncate table syslog
go
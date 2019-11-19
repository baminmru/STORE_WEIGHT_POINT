if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_LOCSIZE]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_LOCSIZE]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_PALLETGOOD]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_PALLETGOOD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_STOCK]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_STOCK]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_STOCK103]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_STOCK103]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_STOCK2]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_STOCK2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_STOCKBLOCKED]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_STOCKBLOCKED]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_STOCK_ALL]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_STOCK_ALL]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_STOCK_pt]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_STOCK_pt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_lOCCHECK_pt]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_lOCCHECK_pt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_bami_vimorozkaGlobal]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_bami_vimorozkaGlobal]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_hranenie]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_hranenie]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_loccheck1]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_loccheck1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_loccheck1_pt]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_loccheck1_pt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_manual]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_manual]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_manual_pt]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_manual_pt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_nedostacha]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_nedostacha]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_palletday]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_palletday]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_prevshipped]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_prevshipped]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_stokmorozdayly]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_stokmorozdayly]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_vimorozka]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_vimorozka]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_vimorozka_blocked]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_vimorozka_blocked]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_vimorozka_rpt]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_vimorozka_rpt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_vimorozka_rpt2]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_vimorozka_rpt2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_vimorozka_rpt3]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_vimorozka_rpt3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_bami_location]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_bami_location]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_LASTPALLETRCV]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_LASTPALLETRCV]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_BAMI_lOCCHECK]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_BAMI_lOCCHECK]
GO




SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create view v_bami_location as
--  ������� ������� ����� � ����� ������ ������������
select id ,code,  loc_qty_pal QMAX, ACTIVE ,left(code,2) sklad, substring(code,4,1) T,  
convert(int,substring(code,4,1))*3 +
convert(int,
case substring(code,5,1)
when 'F' then '1'
when 'N' then '2'
else substring(code,5,1)
end)
 X
,
convert(int,substring(code,7,2))*5+
convert(int,
case substring(code,10,1)
when 'A' then '1'
when 'B' then '2'
when 'C' then '3'
when 'D' then '4'
else substring(code,10,1)
end)
 Y  
,
convert(int,right(code,1)) Z
, right(description,1) LOCTYPE

from location with (nolock)
where code <>'������' 



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE   view v_bami_loccheck1
as
-- ������ ��� �������� ������������ ������ � ������
select distinct A.ID, A.code ,checksum(
ITEM_ID
,LOT_SN
,CUSTOM_FIELD4
,CUSTOM_FIELD6
,CUSTOM_FIELD11
,CUSTOM_FIELD12
,CUSTOM_FIELD7
,item.class
--,month(exp_date),year(exp_date)
) X
 from stock with (nolock)
join v_bami_location A with (nolock)  on A.id = stock.location_id and a.loctype <>'B'
join item on stock.item_id=item.id
join pallet on stock.pallet_id = pallet.id

where pallet_status is null	
group by A.ID, A.code,checksum(
ITEM_ID
,LOT_SN
,CUSTOM_FIELD4
,CUSTOM_FIELD6
,CUSTOM_FIELD11
,CUSTOM_FIELD12
,CUSTOM_FIELD7
,item.class
--,month(exp_date),year(exp_date)
)







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE  view V_BAMI_lOCCHECK as
-- ������ � ������ �������
select code, count(*) CNT from v_bami_loccheck1 with (nolock) group by code
having count(*)>1




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





create view V_BAMI_LASTPALLETRCV as

-- ��� ��������� ������� ����� �� ������ �������
select isnull(max(rec_date),getdate()-1000) LastRCV, pallet
from receiving_history with (nolock)
group by pallet 



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE  view V_BAMI_LOCSIZE as 
select sum(case when p.type='E' then 1 else 1.25 end) cur_qty, l.code, l.loc_qty_pal,plan_loc_qty from 
stock i 
join location l on l.id=i.location_id
join pallet p on i.pallet_id = p.id
where i.pallet_status is null and i.status in (0,100,101,103)
group by l.code, l.loc_qty_pal,plan_loc_qty


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create view V_BAMI_PALLETGOOD as
-- ����������� ������ �� �������� ������
-- ������� �����
select 1 rectype,rec_date the_Date, pallet ,item_id, custom_field4 Factory,custom_field6 country,custom_field11 Kill_place ,custom_field12 isBRAK,
custom_field9 made_date_to, custom_field7 vetsved,
QTY_REC in_Quantity, convert(int,custom_field1) in_Boxes, 
0 out_quantity, 0 out_boxes,
0 stok_quantity, 0 stok_boxes from receiving_history with (nolock)

union all
-- ��������
select 2, ship_date, pallet,item_id,custom_field4,custom_field6,custom_field11,custom_field12,
custom_field9 made_date_to, custom_field7 vetsved,
0,0,QTY_SHIP, convert(int,custom_field1),0,0 Boxes 
from shipping_history with (nolock)

union all
-- ������ �� ������
select 3,getdate(),pallet.code,item_id,custom_field4,custom_field6,custom_field11,custom_field12,
custom_field9 made_date_to, custom_field7 vetsved,
0,0,0,0,QTY_ON_HAND,convert(int,custom_field1) 
from stock  with (nolock)
join pallet with (nolock) on stock.pallet_id = pallet.id  where pallet_status is null



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE       view V_BAMI_STOCK as
-- ��������� ������ � �������������� ���������� ������
select 
stock.id stock_id
, pallet.code pallet_code
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.custom_field9 made_date_to
, stock.custom_field7 vetsved
, stock.LOT_SN Partia
, stock.exp_date
, a.ID LOC_ID
, a.qmax
, a.CODE LOC_CODE
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
, stock.status
, stock.qty_on_hand AtStock
, pallet.type pallettype
 from stock with (nolock)
join item with (nolock) on stock.item_id = item.id
left join  partner with (nolock) on item.class=partner.code
left join pallet with (nolock) on pallet.id = stock.pallet_id
left join v_bami_location A with (nolock)  on A.id = stock.location_id
where pallet_status  is null	and a.active ='Y' and a.loctype <>'B'
and a.code not in (select code  from V_BAMI_LOCCHECK)










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE  view V_BAMI_STOCK103 as
-- ��������� ������ � �������������� ���������� ������
select 
stock.id stock_id
, pallet.code pallet_code
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.QTY_ON_HAND
, stock.exp_date
, a.ID LOC_ID
, a.qmax
, a.CODE LOC_CODE
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
 from stock with (nolock)
join item with (nolock) on stock.item_id = item.id
join  partner with (nolock) on item.class=partner.code
join pallet with (nolock) on pallet.id = stock.pallet_id
join v_bami_location A with (nolock)  on A.id = stock.location_id
where pallet_status  is null and stock.status=103





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO










CREATE        view V_BAMI_STOCK2 as
-- ��������� ������ � �������������� ���������� ������
select 
stock.id stock_id
, pallet.code pallet_code
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.custom_field9 made_date_to
, stock.custom_field7 vetsved
, stock.exp_date
, a.ID LOC_ID
, a.qmax
, a.CODE LOC_CODE
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
, stock.status
, stock.qty_on_hand AtStock
 from stock with (nolock)
join item with (nolock) on stock.item_id = item.id
left join  partner with (nolock) on item.class=partner.code
left join pallet with (nolock) on pallet.id = stock.pallet_id
left join v_bami_location A with (nolock)  on A.id = stock.location_id
where pallet_status  is null	











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE view V_BAMI_STOCKBLOCKED as
-- ��������� ������ � �������������� ���������� ������
select 
stock.id stock_id
, pallet.code pallet_code
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.QTY_ON_HAND
, stock.exp_date
, a.ID LOC_ID
, a.qmax
, a.CODE LOC_CODE
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
 from STOCKBLOCKED stock with (nolock)
join item with (nolock) on stock.item_id = item.id
join  partner with (nolock) on item.class=partner.code
left join pallet with (nolock) on pallet.id = stock.pallet_id
left join v_bami_location A with (nolock)  on A.id = stock.location_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO








create     view V_BAMI_STOCK_ALL as
-- ��������� ������ � �������������� ���������� ������
select 
stock.id stock_id
, pallet.code pallet_code
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.custom_field9 made_date_to
, stock.custom_field7 vetsved
, stock.LOT_SN Partia
, stock.exp_date
, a.ID LOC_ID
, a.qmax
, a.CODE LOC_CODE
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
, stock.status
, stock.qty_on_hand AtStock
, pallet.type pallettype
 from stock with (nolock)
join item with (nolock) on stock.item_id = item.id
left join  partner with (nolock) on item.class=partner.code
left join pallet with (nolock) on pallet.id = stock.pallet_id
left join v_bami_location A with (nolock)  on A.id = stock.location_id
where pallet_status  is null	and a.active ='Y' and a.loctype <>'B'
--and a.code not in (select code  from V_BAMI_LOCCHECK)











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





create   view v_bami_loccheck1_pt
as
-- ������ ��� �������� ������������ ������ � ������
select distinct A.ID, A.code ,checksum(
ITEM_ID
,LOT_SN
,CUSTOM_FIELD4
,CUSTOM_FIELD6
,CUSTOM_FIELD11
,CUSTOM_FIELD12
,CUSTOM_FIELD7
,item.class
,pallet.type
--,month(exp_date),year(exp_date)
) X
 from stock with (nolock)
join v_bami_location A with (nolock)  on A.id = stock.location_id and a.loctype <>'B'
join item on stock.item_id=item.id
join pallet on stock.pallet_id = pallet.id

where pallet_status is null	
group by A.ID, A.code,checksum(
ITEM_ID
,LOT_SN
,CUSTOM_FIELD4
,CUSTOM_FIELD6
,CUSTOM_FIELD11
,CUSTOM_FIELD12
,CUSTOM_FIELD7
,item.class
,pallet.type
--,month(exp_date),year(exp_date)
)


go
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





create   view V_BAMI_lOCCHECK_pt as
-- ������ � ������ �������
select code, count(*) CNT from v_bami_loccheck1_pt with (nolock) group by code
having count(*)>1





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




create view V_BAMI_STOCK_pt as
-- ��������� ������ � �������������� ���������� ������
select 
stock.id stock_id
, pallet.code pallet_code
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.custom_field9 made_date_to
, stock.custom_field7 vetsved
, stock.LOT_SN Partia
, stock.exp_date
, a.ID LOC_ID
, a.qmax
, a.CODE LOC_CODE
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
, stock.status
, stock.qty_on_hand AtStock
, pallet.type pallettype
 from stock with (nolock)
join item with (nolock) on stock.item_id = item.id
left join  partner with (nolock) on item.class=partner.code
left join pallet with (nolock) on pallet.id = stock.pallet_id
left join v_bami_location A with (nolock)  on A.id = stock.location_id
where pallet_status  is null	and a.active ='Y' and a.loctype <>'B'
and a.code not in (select code  from V_BAMI_LOCCHECK_pt)











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create view v_bami_prevshipped
as
-- ���������� �� ��� ������������ ������ (��������� ���� ��������)
select max(rec_date) LastRCV, shipping_history.pallet, ship_date
from receiving_history with (nolock) join
shipping_history with (nolock)
on receiving_history.pallet=shipping_history.pallet and rec_date<=ship_date
group by shipping_history.pallet, ship_date 
--order by shipping_history.pallet, ship_date




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create view V_bami_vimorozkaGlobal
as
-- ������ ���������� ��������� � ������� ������ �������� (item_id)
-- ������� ��������� ����� ������
select a.rectype, b.LastRCV ,a.the_date , a.pallet,a.item_id
,Factory,country,Kill_place ,isBRAK,made_date_to,vetsved
 ,in_quantity , in_boxes
 ,out_quantity, out_boxes
 ,out_quantity * datediff(d, b.LastRCV,a.the_date)/30  +
  stok_quantity * datediff(d, b.LastRCV,getdate())/30  
   dout_quantity
 ,stok_quantity, stok_boxes
from V_BAMI_PALLETGOOD  A with (nolock)
join V_BAMI_LASTPALLETRCV B with (nolock)  on A.pallet=b.pallet and a.the_date >= b.LastRCV

union all

-- ������� ��������� �������� �� ���� ���������� ��������
-- ������� ��� ���������� ������� ������
select 
1 rectype,rec_date LastRCV, rec_date the_Date, 
receiving_history.pallet ,item_id, 
custom_field4 Factory,custom_field6 country,custom_field11 Kill_place ,custom_field12 isBRAK,
custom_field9 made_date_to,custom_field7 vetsved,
QTY_REC in_Quantity, convert(int,custom_field1) in_Boxes, 
0 out_quantity, 0 out_boxes, 0, 0 stok_quantity, 0 stok_boxes 
from receiving_history with (nolock)
join V_BAMI_LASTPALLETRCV B with (nolock)  on receiving_history.pallet=b.pallet and receiving_history.rec_date < b.LastRCV

union all

-- �������� ��� ���������� ������� ������
select 2,v_bami_prevshipped.lastrcv, shipping_history.ship_date, 
shipping_history.pallet,item_id,
custom_field4 Factory,custom_field6 country,custom_field11 Kill_place ,custom_field12 isBRAK,
custom_field9 made_date_to,custom_field7 vetsved,
0,0,
QTY_SHIP, convert(int,custom_field1),
QTY_SHIP * datediff(d, v_bami_prevshipped.LastRCV,shipping_history.ship_date)/30 ,
0,0 Boxes 
from shipping_history with (nolock)
join v_bami_prevshipped with (nolock) on shipping_history.ship_Date = v_bami_prevshipped.ship_date and 
shipping_history.pallet = v_bami_prevshipped.pallet
join V_BAMI_LASTPALLETRCV B with (nolock) on shipping_history.pallet=b.pallet and shipping_history.ship_date < b.LastRCV




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create view v_bami_hranenie
as
-- ��������� ���������������� �� �������������
select 
 partner.code partner_code, 
 sum(in_quantity)  qin
 ,sum(out_quantity) qout
 ,sum( dout_quantity) * 30 /1000 hranenie 
 ,sum(stok_quantity) qstok
from v_bami_vimorozkaGlobal with (nolock)
join item  with (nolock) on item.id = item_id
join partner with (nolock) on item.class = partner.code
group by partner.code



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  view v_bami_manual as
-- ������ �������, ������� �� �������� �������������� �����������
select 
stock.id stock_id
, pallet.code pallet_code
, a.CODE LOC_CODE
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.exp_date
,stock.LOT_SN  Partia
, a.ID LOC_ID
,a.qmax
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
 from stock with (nolock)
join item with (nolock) on stock.item_id = item.id
join  partner with (nolock) on item.class=partner.code
join pallet with (nolock) on pallet.id = stock.pallet_id
join v_bami_location A with (nolock)  on A.id = stock.location_id
where pallet_status  is null	and a.active ='Y' and a.loctype <>'B'
and a.code  in (select code  from V_BAMI_LOCCHECK with (nolock))






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





create   view v_bami_manual_pt as
-- ������ �������, ������� �� �������� �������������� �����������
select 
stock.id stock_id
, pallet.code pallet_code
, a.CODE LOC_CODE
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.exp_date
,stock.LOT_SN  Partia
, a.ID LOC_ID
,a.qmax
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
 from stock with (nolock)
join item with (nolock) on stock.item_id = item.id
join  partner with (nolock) on item.class=partner.code
join pallet with (nolock) on pallet.id = stock.pallet_id
join v_bami_location A with (nolock)  on A.id = stock.location_id
where pallet_status  is null	and a.active ='Y' and a.loctype <>'B'
and a.code  in (select code  from V_BAMI_LOCCHECK_pt with (nolock))







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create view v_bami_nedostacha
as
-- ����� ������������ �������� ��� ���������� �� ������� � ������� ���� ������ (item_id)
select 
 item_id, item.code, item.description,
 sum(in_quantity)  qin
 ,sum(out_quantity) qout
 ,sum( dout_quantity)*0.0005 vimorozka
 ,sum(out_quantity)*0.001+sum(in_quantity)*0.001 pogreshnost
 ,sum(stok_quantity) qstok
 ,sum(in_quantity) -sum(out_quantity) mustbeinstok
 ,-(sum(in_quantity) -sum(out_quantity)-sum(stok_quantity) -sum( dout_quantity)*0.0005 -sum(out_quantity)*0.001) nedostacha
from v_bami_vimorozkaGlobal with (nolock)
join item with (nolock) on item.id = item_id
group by item_id,item.code, item.description
having sum(in_quantity) >0  and sum(in_quantity) -sum(out_quantity)-sum(stok_quantity) -sum( dout_quantity)*0.0005 -sum(out_quantity)*0.001 <0 or
sum(in_quantity) -sum(out_quantity)-sum(stok_quantity) > 0

--order by item.code




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create view v_bami_palletday as

 

-- ������� ��������� ����� ������

select a.the_date , a.pallet,a.item_id,

datediff(d, b.LastRCV,a.the_date)as  days

from V_BAMI_PALLETGOOD  A with (nolock)

join V_BAMI_LASTPALLETRCV B with (nolock)  on A.pallet=b.pallet and a.the_date >= b.LastRCV

 

union all

-- �������� ��� ���������� ������� ������

select  shipping_history.ship_date, 

shipping_history.pallet, shipping_history.item_id,

datediff(d, v_bami_prevshipped.LastRCV,shipping_history.ship_date)

from shipping_history with (nolock)

join v_bami_prevshipped with (nolock) on shipping_history.ship_Date = v_bami_prevshipped.ship_date and 

shipping_history.pallet = v_bami_prevshipped.pallet

join V_BAMI_LASTPALLETRCV B with (nolock) on shipping_history.pallet=b.pallet and shipping_history.ship_date < b.LastRCV


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create view v_bami_stokmorozdayly as
-- ������ ��� ����������� ��������� �� ������� ���
select 
 partner.code partner_code ,item_id, item.code item_code, item.description,
 0  qin
 ,0 qout
 ,sum( QTY_ON_HAND)* 1/30*0.0005 vimorozka
 ,0 pogreshnost
 ,sum(QTY_ON_HAND) qstok
 ,sum(QTY_ON_HAND) mustbeinstok
 ,0 nedostacha
from stock  with (nolock)
join pallet  with (nolock)on stock.pallet_id = pallet.id  and stock.pallet_status is null
join item  with (nolock)on item.id = item_id
join partner  with (nolock)on item.class = partner.code
group by partner.code, item_id,item.code, item.description



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create view v_bami_vimorozka
as
-- ��������� � ������� ������� �� ��������� ������ ������
select a.rectype, b.LastRCV ,a.the_date , a.pallet,a.item_id
 ,in_quantity , in_boxes
 ,out_quantity, out_boxes
 ,out_quantity * datediff(d, b.LastRCV,a.the_date)/30  
  + stok_quantity * datediff(d, b.LastRCV,a.the_date)/30
 dout_quantity
 ,stok_quantity, stok_boxes
from V_BAMI_PALLETGOOD  A with (nolock)
join V_BAMI_LASTPALLETRCV B with (nolock)  on A.pallet=b.pallet and a.the_date >= b.LastRCV
--order by a.pallet, rectype



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE   view v_bami_vimorozka_blocked as 
select 
 partner.code partner_code ,item_id, item.code item_code, item.description,
 Factory,country,Kill_place ,isBRAK,made_date_to,vetsved,
  in_quantity 
 ,out_quantity 
 ,dout_quantity
 ,stok_quantity 
from v_bami_vimorozkaGlobal with (nolock)
join item  with (nolock)on item.id = item_id
join partner  with (nolock)on item.class = partner.code

union all 

-- ��������������� �� ��������� �������
select 
 partner.code partner_code ,item_id, item.code item_code, item.description,
custom_field4 Factory,custom_field6 country,
custom_field11 Kill_place ,custom_field12 isBRAK,
custom_field9 made_date_to,custom_field7 vetsved,
0,0,-QTY_ON_HAND * 2000 ,0
from stockblocked with (nolock)
left join pallet on pallet.id = stockblocked.pallet_id
join item  with (nolock)on item.id = item_id
join partner  with (nolock)on item.class = partner.code






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create view v_bami_vimorozka_rpt as

select 
 partner.code partner_code ,item_id, item.code item_code, item.description,
 sum(in_quantity)  qin
 ,sum(out_quantity) qout
 ,sum( dout_quantity)*0.0005 vimorozka
 ,sum(out_quantity)*0.001 +sum(in_quantity)*0.001 pogreshnost
 ,sum(stok_quantity) qstok
 ,sum(in_quantity) -sum(out_quantity) mustbeinstok
 ,-(sum(in_quantity) -sum(out_quantity)-sum(stok_quantity) -sum( dout_quantity)*0.0005  -sum(out_quantity)*0.001) nedostacha
from v_bami_vimorozkaGlobal with (nolock)
join item  with (nolock)on item.id = item_id
join partner  with (nolock)on item.class = partner.code
group by partner.code, item_id,item.code, item.description
 


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE     view v_bami_vimorozka_rpt2 as

select 
 partner.code partner_code ,item_id, item.code item_code, item.description,
 Factory,country,Kill_place ,isBRAK,made_date_to,vetsved,
 sum(in_quantity)  qin
 ,sum(out_quantity) qout
 ,sum( dout_quantity)*0.0005 vimorozka
 ,sum(out_quantity)*0.001 +sum(in_quantity)*0.001 +sum(stok_quantity)*0.001 pogreshnost
 ,sum(stok_quantity) qstok
 ,-( 
    sum(in_quantity) - sum(out_quantity)- sum(stok_quantity) 
    - ( sum( dout_quantity)*0.0005+sum(out_quantity)*0.001 +sum(in_quantity)*0.001 +sum(stok_quantity)*0.001)
    ) otbor
, sum(stok_quantity)+( 
    sum(in_quantity) - sum(out_quantity)- sum(stok_quantity) 
    - ( sum( dout_quantity)*0.0005+sum(out_quantity)*0.001 +sum(in_quantity)*0.001 +sum(stok_quantity)*0.001)
    ) to_ship
from v_bami_vimorozka_blocked with (nolock)
join item  with (nolock)on item.id = item_id
join partner  with (nolock)on item.class = partner.code
group by partner.code, item_id,item.code, item.description,Factory,country,Kill_place ,isBRAK,made_date_to,vetsved
 
having sum(stok_quantity)+( 
    sum(in_quantity) - sum(out_quantity)- sum(stok_quantity) 
    - ( sum( dout_quantity)*0.0005+sum(out_quantity)*0.001 +sum(in_quantity)*0.001 +sum(stok_quantity)*0.001)
    ) >0
and 
-( 
    sum(in_quantity) - sum(out_quantity)- sum(stok_quantity) 
    - ( sum( dout_quantity)*0.0005+sum(out_quantity)*0.001 +sum(in_quantity)*0.001 +sum(stok_quantity)*0.001)
    )>0







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create     view v_bami_vimorozka_rpt3 as

select 
 partner.code partner_code ,item_id, item.code item_code, item.description,
 Factory,country,Kill_place ,isBRAK,made_date_to,vetsved,
 sum(in_quantity)  qin
 ,sum(out_quantity) qout
 ,sum( dout_quantity)*0.0005 vimorozka
 ,sum(out_quantity)*0.001 +sum(in_quantity)*0.001 +sum(stok_quantity)*0.001 pogreshnost
 ,sum(stok_quantity) qstok
 ,-( 
    sum(in_quantity) - sum(out_quantity)- sum(stok_quantity) 
    - ( sum( dout_quantity)*0.0005+sum(out_quantity)*0.001 +sum(in_quantity)*0.001 +sum(stok_quantity)*0.001)
    ) otbor
, sum(stok_quantity)+( 
    sum(in_quantity) - sum(out_quantity)- sum(stok_quantity) 
    - ( sum( dout_quantity)*0.0005+sum(out_quantity)*0.001 +sum(in_quantity)*0.001 +sum(stok_quantity)*0.001)
    ) to_ship
from v_bami_vimorozka_blocked with (nolock)
join item  with (nolock)on item.id = item_id
join partner  with (nolock)on item.class = partner.code
group by partner.code, item_id,item.code, item.description,Factory,country,Kill_place ,isBRAK,made_date_to,vetsved
 


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

alter table pallet add TYPE varchar(40) null
go
update pallet set TYPE = 'E' where type is null
go


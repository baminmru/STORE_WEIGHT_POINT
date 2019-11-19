



drop view v_bami_location
go
create view v_bami_location as
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

from location 
where code <>'Пустая' 
go


drop view
v_bami_loccheck1
go

create view v_bami_loccheck1
as

select distinct A.ID, code,checksum(
ITEM_ID

--,CUSTOM_FIELD4
--,CUSTOM_FIELD6
--,CUSTOM_FIELD11
,CUSTOM_FIELD12
--,month(exp_date),year(exp_date)
) X
 from stock 
join v_bami_location A on A.id = stock.location_id and a.loctype <>'B'
where pallet_status is null	
group by A.ID, code,checksum(
ITEM_ID
--,CUSTOM_FIELD4
--,CUSTOM_FIELD6
--,CUSTOM_FIELD11
,CUSTOM_FIELD12
--,month(exp_date),year(exp_date)
)

go 

drop view V_BAMI_lOCCHECK
go
create view V_BAMI_lOCCHECK as
select code, count(*) CNT from v_bami_loccheck1 group by code
having count(*)>1
go




drop view   V_BAMI_STOCK
go
create view V_BAMI_STOCK as
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
, stock.LOT_SN Partia
 from stock with (nolock)
join item with (nolock) on stock.item_id = item.id
left join  partner with (nolock) on item.class=partner.code
left join pallet with (nolock) on pallet.id = stock.pallet_id
left join v_bami_location A with (nolock)  on A.id = stock.location_id
where pallet_status  is null	and a.active ='Y' and a.loctype <>'B'
and a.code not in (select code  from V_BAMI_LOCCHECK)


go

drop view v_bami_manual
go
create view v_bami_manual as
select 
stock.id stock_id
, pallet.code pallet_code
, a.CODE LOC_CODE
,partner.code Partner,item.code item_code, item.description , stock.custom_field6 Country
, stock.custom_field4 Factory
, stock.custom_field11 KILL_PLACE
, stock.custom_field12 IsBRAK
, stock.exp_date
, a.ID LOC_ID
,a.qmax
, a.sklad
, a.T
, a.X
, a.Y
, a.Z
 from stock 
join item on stock.item_id = item.id
join  partner on item.class=partner.code
join pallet on pallet.id = stock.pallet_id
join v_bami_location A on A.id = stock.location_id
where pallet_status is null	and a.active ='Y' and a.loctype <>'B'
and a.code  in (select code  from V_BAMI_LOCCHECK)

go

drop view V_BAMI_PALLETGOOD
go

create view V_BAMI_PALLETGOOD as
select 1 rectype,rec_date the_Date,    pallet ,item_id, custom_field4 Factory,custom_field6 country,custom_field11 Kill_place ,custom_field12 isBRAK,
QTY_REC in_Quantity, convert(int,custom_field1) in_Boxes, 
0 out_quantity, 0 out_boxes,
 0 stok_quantity, 0 stok_boxes from receiving_history 
union all
select 2, ship_date, pallet,item_id,custom_field4,custom_field6,custom_field11,custom_field12,0,0,QTY_SHIP, convert(int,custom_field1),0,0 Boxes from shipping_history 
union all
select 3,getdate(),pallet.code,item_id,custom_field4,custom_field6,custom_field11,custom_field12,0,0,0,0,QTY_ON_HAND,convert(int,custom_field1) from stock  
join pallet on stock.pallet_id = pallet.id  where pallet_status is null
go


drop view V_BAMI_LASTPALLETRCV
go


create view V_BAMI_LASTPALLETRCV as
select isnull(max(rec_date),getdate()-1000) LastRCV, pallet
from receiving_history 
group by pallet 
go


drop view v_bami_prevshipped
go 
create view v_bami_prevshipped
as
select max(rec_date) LastRCV, shipping_history.pallet, ship_date
from receiving_history join
shipping_history 
on receiving_history.pallet=shipping_history.pallet and rec_date<=ship_date
group by shipping_history.pallet, ship_date 
--order by shipping_history.pallet, ship_date

go


drop view v_bami_vimorozka 
go

create view v_bami_vimorozka
as
select a.rectype, b.LastRCV ,a.the_date , a.pallet,a.item_id
 ,in_quantity , in_boxes
 ,out_quantity, out_boxes
 ,out_quantity * datediff(d, b.LastRCV,a.the_date)/30  
  + stok_quantity * datediff(d, b.LastRCV,a.the_date)/30
 dout_quantity
 ,stok_quantity, stok_boxes
from V_BAMI_PALLETGOOD A
join V_BAMI_LASTPALLETRCV B on A.pallet=b.pallet and a.the_date >= b.LastRCV
--order by a.pallet, rectype
go




drop view  V_bami_vimorozkaGlobal
go
create view V_bami_vimorozkaGlobal
as
select a.rectype, b.LastRCV ,a.the_date , a.pallet,a.item_id
 ,in_quantity , in_boxes
 ,out_quantity, out_boxes
 ,out_quantity * datediff(d, b.LastRCV,a.the_date)/30  +
  stok_quantity * datediff(d, b.LastRCV,getdate())/30  
   dout_quantity
 ,stok_quantity, stok_boxes
from V_BAMI_PALLETGOOD A
join V_BAMI_LASTPALLETRCV B on A.pallet=b.pallet and a.the_date >= b.LastRCV

union all

select 
1 rectype,rec_date LastRCV, rec_date the_Date, 
receiving_history.pallet ,item_id, 
   --custom_field4 Factory,custom_field6 country,custom_field11 Kill_place ,custom_field12 isBRAK,
QTY_REC in_Quantity, convert(int,custom_field1) in_Boxes, 
0 out_quantity, 0 out_boxes, 0, 0 stok_quantity, 0 stok_boxes 
from receiving_history 
join V_BAMI_LASTPALLETRCV B on receiving_history.pallet=b.pallet and receiving_history.rec_date < b.LastRCV
union all
select 2,v_bami_prevshipped.lastrcv, shipping_history.ship_date, 
shipping_history.pallet,item_id,
--custom_field4,custom_field6,custom_field11,custom_field12,
0,0,
QTY_SHIP, convert(int,custom_field1),
QTY_SHIP * datediff(d, v_bami_prevshipped.LastRCV,shipping_history.ship_date)/30 ,
0,0 Boxes 
from shipping_history 
join v_bami_prevshipped on shipping_history.ship_Date = v_bami_prevshipped.ship_date and 
shipping_history.pallet = v_bami_prevshipped.pallet
join V_BAMI_LASTPALLETRCV B on shipping_history.pallet=b.pallet and shipping_history.ship_date < b.LastRCV

go

drop view v_bami_vimorozka_rpt
go

create view v_bami_vimorozka_rpt as

select 
 partner.code partner_code ,item_id, item.code item_code, item.description,
 sum(in_quantity)  qin
 ,sum(out_quantity) qout
 ,sum( dout_quantity)*0.0005 vimorozka
 ,sum(out_quantity)*0.001 pogreshnost
 ,sum(stok_quantity) qstok
 ,sum(in_quantity) -sum(out_quantity) mustbeinstok
 ,-(sum(in_quantity) -sum(out_quantity)-sum(stok_quantity) -sum( dout_quantity)*0.0005  -sum(out_quantity)*0.001) nedostacha
from v_bami_vimorozkaGlobal
join item on item.id = item_id
join partner on item.class = partner.code
group by partner.code, item_id,item.code, item.description
 go

drop view v_bami_nedostacha
go

create view v_bami_nedostacha
as

select 
 item_id, item.code, item.description,
 sum(in_quantity)  qin
 ,sum(out_quantity) qout
 ,sum( dout_quantity)*0.0005 vimorozka
 ,sum(out_quantity)*0.001 pogreshnost
 ,sum(stok_quantity) qstok
 ,sum(in_quantity) -sum(out_quantity) mustbeinstok
 ,-(sum(in_quantity) -sum(out_quantity)-sum(stok_quantity) -sum( dout_quantity)*0.0005 -sum(out_quantity)*0.001) nedostacha
from v_bami_vimorozkaGlobal
join item on item.id = item_id
group by item_id,item.code, item.description
having sum(in_quantity) >0  and sum(in_quantity) -sum(out_quantity)-sum(stok_quantity) -sum( dout_quantity)*0.0005 -sum(out_quantity)*0.001 <0 or
sum(in_quantity) -sum(out_quantity)-sum(stok_quantity) > 0

--order by item.code

go


drop view v_bami_hranenie
go 
create view v_bami_hranenie
as
select 
 partner.code partner_code, 
 sum(in_quantity)  qin
 ,sum(out_quantity) qout
 ,sum( dout_quantity) * 30 /1000 hranenie 
 ,sum(stok_quantity) qstok
from v_bami_vimorozkaGlobal
join item on item.id = item_id
join partner on item.class = partner.code
group by partner.code
go




create view v_bami_stokmorozdayly as
select 
 partner.code partner_code ,item_id, item.code item_code, item.description,
 0  qin
 ,0 qout
 ,sum( QTY_ON_HAND)* 1/30*0.0005 vimorozka
 ,0 pogreshnost
 ,0 qstok
 ,0 mustbeinstok
 ,0 nedostacha
from stock  
join pallet on stock.pallet_id = pallet.id  and stock.pallet_status is null
join item on item.id = item_id
join partner on item.class = partner.code
group by partner.code, item_id,item.code, item.description
go


alter table pallet add TYPE varchar(40) null
go
update pallet set TYPE = 'E'
go

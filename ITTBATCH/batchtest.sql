delete from palmtest..placetab

go

declare @id int
declare @p varchar(255)
declare @l varchar(255)

set @id =1
declare lc  cursor for select top 1000 pallet.code,location.code from pallet, location 
open lc

fetch next from lc into @p,@l
while @@fetch_status >=0 
begin
       set @id = @id+1
       insert into palmtest..placetab(uniqueid,pallet,location,dirty) values(@id,@p,@l,0)
       fetch next from lc into @p,@l
end
close lc
deallocate lc


go
delete from palmtest..loadtab

go

declare @id int
declare @cnt  int
declare @p varchar(255)
declare @l varchar(255)
set @cnt =1
set @id =36000
declare @q_num varchar(255)

set @q_num ='80043874 от 14/08/2007'

while @cnt <20
begin
       set @id = @id+1
       insert into palmtest..loadtab(uniqueid,qrynum, pallet,ssctop, sscbottom,dirty) values(@id,@q_num,@id,'020500018607900037010890121212','00350001861030006654130105259301023648',0)
       set @cnt=@cnt+1	
end

/*

   select top 1 convert(varchar(20),number) + ' от '+ convert(varchar(20),rec_date,103) from RECEIVING_ORDER
   join partner on RECEIVING_ORDER.partner_id=partner.id
   Where (Status = 0 Or Status = 1) and rec_date > getdate()-15 and partner.code ='Unilever'
*/  

--select * from item
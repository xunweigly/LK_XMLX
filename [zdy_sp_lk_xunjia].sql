USE [UFDATA_999_2017]
GO
/****** Object:  StoredProcedure [dbo].[zdy_sp_lk_xunjia]    Script Date: 12/28/2017 15:39:56 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER proc [dbo].[zdy_sp_lk_xunjia](@cno varchar(20),@error nvarchar(100) output)
as
set nocount on 
begin
declare @pk int ,@cnt int,@max int

SELECT @cnt = 0

select @pk = LK1_0007_E001_PK from LK_XM_LX  WHERE cNo = @cno

---根据编码，名称，规格判断是否已经下询价单
select @cnt = COUNT(*) from 
	(select cinvcode,cinvname,cinvstd,SUM(iquantity) zsl  from LK1_XM_BOM
	where LK1_0007_E001_PK=@pk group by cinvcode,cinvname,cinvstd)  xm
where Not exists (select 1 from U8CUSTDEF_0015_E001 a ,U8CUSTDEF_0015_E002 b
where a.U8CUSTDEF_0015_E001_PK =b.U8CUSTDEF_0015_E001_PK
and a.citemname= @cno and b.cinvcode = xm.cinvcode and b.cinvstd = xm.cinvstd
and b.cinvname=xm.cinvname)


if @cnt = 0 OR @cnt  IS NULL
begin
	select @error='物料已经全部下达询价，禁止重复下达'
	return

end

declare @id int,@ids int
--根据最大行号回写
select @id=maxvalue from UAP_TablePKIntMaxValue where tablename='U8CUSTDEF_0015_E001'
update UAP_TablePKIntMaxValue set maxvalue=maxvalue+@cnt  where tablename='U8CUSTDEF_0015_E001'

select @ids=maxvalue from UAP_TablePKIntMaxValue where tablename='U8CUSTDEF_0015_E002'
update UAP_TablePKIntMaxValue set maxvalue=maxvalue+@cnt  where tablename='U8CUSTDEF_0015_E002'


declare @rowno int
select @rowno=1
declare @cinvcode varchar(100),@cinvname varchar(100),@cinvstd varchar(100),@jldw varchar(100)
declare @zsl decimal(18,3) 


declare ybyh cursor for
select xm.cinvcode,xm.cinvname,xm.cinvstd,jiliangdw,xm.zsl from
(select cinvcode,cinvname,cinvstd,jiliangdw,SUM(iquantity) zsl  from LK1_XM_BOM
	where LK1_0007_E001_PK=@pk group by cinvcode,cinvname,cinvstd,jiliangdw)  xm
where Not exists (select 1 from U8CUSTDEF_0015_E001 a ,U8CUSTDEF_0015_E002 b
where a.U8CUSTDEF_0015_E001_PK =b.U8CUSTDEF_0015_E001_PK
and a.citemname= @cno and b.cinvcode = xm.cinvcode and b.cinvstd = xm.cinvstd
and b.cinvname=xm.cinvname)
open ybyh

-- 打开游标
--给参数赋值
fetch next from ybyh into @cinvcode,@cinvname,@cinvstd,@jldw,@zsl
--执行游标第一条记录
WHILE @@FETCH_STATUS = 0
BEGIN

---主键用id+行号

insert into U8CUSTDEF_0015_E001(cNo,cMaker,U8CUSTDEF_0015_E001_PK,UAPRuntime_RowNO,citemname,cbz,ccusname,ddate,dmakedate,iswfcontrolled,bcgxj,BRE)
select 'XM'+CAST(@id+@rowno AS VARCHAR(10)),cmaker,@id+@rowno,1,cNo,'项目询价','项目部',GETDATE(),GETDATE(),1,1,1 from LK_XM_LX 
where cNo=@cno



insert into U8CUSTDEF_0015_E002(U8CUSTDEF_0015_E001_PK,U8CUSTDEF_0015_E002_PK,UAPRuntime_RowNO,cinvcode,cinvname,cinvstd,cqty1)
select @id+@rowno,@ids+@rowno,1,@cinvcode,@cinvname,@cinvstd,CAST(@zsl AS VARCHAR(20))+@jldw from LK_XM_LX   
where cNo=@cno


select @rowno = @rowno+1
fetch next from ybyh into @cinvcode,@cinvname,@cinvstd,@jldw,@zsl
END

close ybyh -- 关闭游标
deallocate ybyh -- 删除游标




end

/****** Object:  Trigger [dbo].[zdy_tr_lkxmztgg]    Script Date: 01/16/2019 15:19:42 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER TRIGGER [dbo].[zdy_tr_lkxmztgg]
   ON  [dbo].[LK_XM_LX]
   AFTER UPDATE
AS 
BEGIN
	SET NOCOUNT ON;
	--获得主键id
	declare @id int,@ckcnt int
	declare @xmcno varchar(20),@cmaker varchar(20)
	DECLARE @xmzt VARCHAR(20),@cauditor VARCHAR(20)
	DECLARE @xmzt2 VARCHAR(20),@cauditor2 VARCHAR(20)

	select @xmcno =Inserted.cNo ,@xmzt = Inserted.xmzt,@cauditor = Inserted.cAuditor FROM inserted 
	select @xmzt2 = xmzt ,@cauditor2  =cAuditor fROM Deleted 
		
if update(iverifystate) OR UPDATE(cAuditor)
begin
 declare @istatein int,@istatedel int
 select @istatein = iverifystate  from inserted
  select @istatedel = iverifystate  from deleted
  --提交
  if (isnull(@istatein ,0)=1 and ISNULL(@istatedel,0)=0)
  begin
  update LK_XM_LX set xmzt='审核中' from LK_XM_LX a,inserted b where a.LK1_0007_E001_PK = b.LK1_0007_E001_PK

  end
  else if isnull(@istatein ,0)=2 and isnull(@istatedel ,0)<>2
  begin
  update LK_XM_LX set xmzt='成功' from LK_XM_LX a,inserted b where a.LK1_0007_E001_PK = b.LK1_0007_E001_PK
  EXEC dbo.zdy_lk_sp_Item_Insert  @cItemName = @xmcno, @Ccode = N'FN' 
  end
  else if (ISNULL(@istatein ,0)<>2 and isnull(@istatedel ,0)=2) OR (  ISNULL(@cauditor,'')='' AND ISNULL(@cauditor2,'')<>'')
  begin
		IF NOT  EXISTS(SELECT 1 FROM dbo.PU_AppVouchs WHERE LEFT(cItemCode,LEN(cItemCode)-2) = @xmcno) AND  NOT  EXISTS(SELECT 1 FROM dbo.MaterialAppVouchs WHERE LEFT(cItemCode,LEN(cItemCode)-2) = @xmcno) AND  NOT  EXISTS(SELECT 1 FROM dbo.rdrecords11 WHERE LEFT(cItemCode,LEN(cItemCode)-2) = @xmcno) AND  NOT  EXISTS(SELECT 1 FROM dbo.rdrecords10 WHERE LEFT(cItemCode,LEN(cItemCode)-2) = @xmcno)
		BEGIN
		delete from fitemss97 from fitemss97,fitemss97class where   fitemss97.cItemCcode = fitemss97class.cItemCcode and  fitemss97class.citemcname =@xmcno
	delete from fitemss97class where citemcname =	@xmcno	
		end	
		ELSE
	UPDATE fitemss97 SET bclose=1	from fitemss97,fitemss97class where   fitemss97.cItemCcode = fitemss97class.cItemCcode and  fitemss97class.citemcname =@xmcno

		end
  
END


END

create TRIGGER [dbo].[zdy_tr_lkxmztdel]
   ON  [dbo].[LK_XM_LX]
   AFTER delete
AS 
BEGIN
	SET NOCOUNT ON;
	--»ñµÃÖ÷¼üid
	declare @id int,@ckcnt int
	declare @xmcno varchar(20),@cmaker varchar(20)
	DECLARE @xmzt2 VARCHAR(20),@xmwczt2 VARCHAR(20)
	
	select @xmcno =cNo,@xmzt2 = xmzt,@xmwczt2 = xmwczt  FROM Deleted 

	
		
		IF NOT  EXISTS(SELECT 1 FROM dbo.PU_AppVouchs WHERE LEFT(cItemCode,LEN(cItemCode)-2) = @xmcno) AND  NOT  EXISTS(SELECT 1 FROM dbo.MaterialAppVouchs WHERE LEFT(cItemCode,LEN(cItemCode)-2) = @xmcno) AND  NOT  EXISTS(SELECT 1 FROM dbo.rdrecords11 WHERE LEFT(cItemCode,LEN(cItemCode)-2) = @xmcno) AND  NOT  EXISTS(SELECT 1 FROM dbo.rdrecords10 WHERE LEFT(cItemCode,LEN(cItemCode)-2) = @xmcno)
		BEGIN
		delete from fitemss97 from fitemss97,fitemss97class where   fitemss97.cItemCcode = fitemss97class.cItemCcode and  fitemss97class.citemcname =@xmcno
	delete from fitemss97class where citemcname =	@xmcno	
		end	
		ELSE
	UPDATE fitemss97 SET bclose=1	from fitemss97,fitemss97class where   fitemss97.cItemCcode = fitemss97class.cItemCcode and  fitemss97class.citemcname =@xmcno
			
		end
		

		

	
   


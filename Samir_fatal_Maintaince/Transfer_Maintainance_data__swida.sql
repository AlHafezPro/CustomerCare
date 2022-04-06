--- «·«—ﬁ«„ «·„Œ“‰Ì…
-- select * from [192.168.221.1].newhalls2010.dbo.pieces
delete from [192.168.221.1].newhalls2010.dbo.pieces
Insert Into [192.168.221.1].newhalls2010.dbo.Pieces (PieceNo, PieceName, PieceStockNo, qty, CliPrice, DistPrice, DealPrice, UniteNo, AccNo, Notes, famNo, PAccNo, RpAccNo) 
Select PieceNo, PieceName, PieceStockNo, qty, CliPrice, DistPrice, DealPrice, UniteNo, AccNo, Notes, famNo, PAccNo, RpAccNo From Hafez2010.dbo.pieces

--- «·„‰ Ã« 
delete from [192.168.221.1].newhalls2010.dbo.adhammodels
insert into [192.168.221.1].newhalls2010.dbo.adhammodels(ModNo, AccNo, RetAccNo, Symbol, Name, GrpNo, FamNo, DealPrice, DistPrice, ConsPrice, DealDisc, DistDisc, ModYear, ProdKind, InventPoint, ItemNo) 
select ModNo, AccNo, RetAccNo, Symbol, Name, GrpNo, FamNo, DealPrice, DistPrice, ConsPrice, DealDisc, DistDisc, ModYear, ProdKind, InventPoint, ItemNo From Hafez2010.dbo.adhammodels

--- «·⁄«∆·« 
delete from [192.168.221.1].newhalls2010.dbo.adhamproductfamily
insert into [192.168.221.1].newhalls2010.dbo.adhamproductfamily(ProdFamNo, ProdFamName, ProdFamNameA, ProdFamOrd)
Select ProdFamNo, ProdFamName, ProdFamNameA, ProdFamOrd from Hafez2010.dbo.adhamproductfamily



--- «·√⁄ÿ«· 
-- Delete from  [192.168.221.1].newhalls2010.dbo.ReparationWorks
-- Insert into [192.168.221.1].newhalls2010.dbo.ReparationWorks(Id_ComNo, RepNo, RepTypeNo, Price, Notes) 
-- Select Id_ComNo, RepNo, RepTypeNo, Price, Notes from Hafez2010.dbo.ReparationWorks
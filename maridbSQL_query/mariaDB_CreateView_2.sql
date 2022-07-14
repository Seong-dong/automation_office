-- USE {DatabaseName};
/*
CREATE VIEW item_code_view AS
SELECT CONCAT(TYPE, LPAD(IDX, "4", "0")) AS "CODE", AA.NAME, AA.SPEC, AA.INGREDIENT, AA.UNIT, DATE_FORMAT(AA.REGDATE, '%Y-%m-%d') AS REGDATE, AA.REGISTANT, AA.STYLENAME FROM (
select * from ecount_item_code
UNION all
select * from ecount_item_market
) AS AA;
*/
-- item_code_view
-- DROP VIEW item_code_view;
-- SELECT * FROM VIEWS;

/*
ALTER TABLE AD_POS_TITLE_INFORMATION AUTO_INCREMENT=1;
SET @COUNT = 0;
UPDATE AD_POS_TITLE_INFORMATION SET ID = @COUNT:=@COUNT+1;
*/

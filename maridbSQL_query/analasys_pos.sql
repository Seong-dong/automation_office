USE sdjo;

SET @idx:=0;

SELECT distinct
	@idx:=@idx+1 AS idx,
	doc_ship.`Ex-Fac` AS 'Shp_ExFac',
	doc_ship.BRAND_CODE AS 'Shp_BC', 
	doc_ship.FACTORY AS 'Shp_FAC', 
	doc_ship.`INVOICE#` AS 'Shp_INVOICE No', 
	doc_ship.`PO#` AS 'Shp_PO No',
	doc_pos.`PO No` AS 'POS_PO No'
FROM doc_factory_ship AS doc_ship
LEFT JOIN mdPosMergeTable AS doc_pos
ON doc_ship.`PO#` = doc_pos.`PO No`
WHERE BRAND_CODE != "";

-- SELECT * FROM mdPosMergeTable;
USE sdjo;

-- UPDATE ecount_item_code SET UNIT = 'YDS' WHERE NAME LIKE '%FABRIC%';
-- SELECT * FROM ecount_item_code WHERE NAME LIKE '%FABRIC%';
-- SELECT *, 'f' AS brand FROM mdPosMergeTable;
/*
CREATE VIEW saOrder_view AS
SELECT
	Brand,
	DATE_FORMAT(`R/D`,'%Y-%m-%d') AS `Regist Date`,
	CONCAT(Season, ' | ', `WS No`, ' | ', `Style Code`, ' | ', `Style Name`, ' | ', `Color Code`, ' | ', `Color Name`, ' | ', Size, ' | ', Factory, ' | ', Destination) AS ID, 
	`Total Qty` AS Qty,
	AMOUNT AS Amount,
	Currency
FROM saPosMergeTable
WHERE `WS No` IS NOT NULL;

*/
/*
CREATE VIEW saDelivery_view AS
SELECT
	Brand,
	DATE_FORMAT(`Date`,'%Y-%m-%d') AS `Ship Date`,
	CONCAT(Season, ' | ', `WS No`, ' | ', `Style Code`, ' | ', `Style Name`, ' | ', `Color Code`, ' | ', `Color Name`, ' | ', Size, ' | ', Factory, ' | ', Destination) AS ID, 
	Currency,
	`Unit Price` AS Price,
	Qty,
	`Unit Price` * Qty AS Amount   
FROM saDelMergeTable
WHERE `WS No` IS NOT NULL;
*/
/*
CREATE VIEW saInv_view AS
SELECT 
	Brand,
	DATE_FORMAT(`Date`,'%Y-%m-%d') AS `Tax Date`,
	CONCAT(Season, ' | ', `WS No`, ' | ', `Style Code`, ' | ', `Style Name`, ' | ', `Color Code`, ' | ', `Color Name`, ' | ', Size, ' | ', Factory, ' | ', Destination) AS ID, 
	Currency,
	`Unit Price` AS Price,
	Qty,
	`Unit Price` * Qty AS Amount   
FROM saInvMergeTable
WHERE `WS No` IS NOT NULL;
*/
-- SELECT * FROM saOrder_view WHERE Brand = 'AJO' AND ID = '21FW | AJO-001 | 21FWCAP01 | - | - | BLACK | 57CM | PVT | KOREA';
-- SELECT * FROM saDelivery_view WHERE Brand = 'AJO';
/*
SET @idx:=0;
@idx:=@idx+1 AS idx,	*/

-- 오더리스트랑 납품리스트 비교
/*
CREATE VIEW saOrderListCompareToDelivery_view AS
SELECT A.Brand, A.`Regist Date` AS RD, A.ID AS OD_ID, SUM(A.Qty) AS order_total_qty , B.ID AS DEL_ID, SUM(B.Qty) AS delivery_total_qty
FROM saOrder_view AS A
LEFT JOIN (SELECT C.Brand, C.`Ship Date`, C.ID, C.Currency, SUM(C.Qty) AS Qty, SUM(C.Amount) AS Amount FROM saDelivery_view AS C
GROUP BY ID) AS B
ON A.ID = B.ID
-- WHERE A.`Regist Date` BETWEEN '2021-07-01' AND '2022-12-31' -- AND A.Brand = 'DYN'
GROUP BY A.ID;
*/
-- ORDER BY B.ID ASC

/*
SELECT  A.Brand, A.`Regist Date` AS RD, A.ID, SUM(A.Qty), B.ID, SUM(B.Qty)
FROM saOrder_view AS A
LEFT JOIN saDelivery_view AS B
ON A.ID = B.ID
WHERE A.`Regist Date` BETWEEN '2021-07-01' AND '2022-12-31' AND A.Brand = 'DYN'
GROUP BY A.ID
ORDER BY B.ID ASC
*/
/*
SELECT *, SUM(ID) 
FROM saOrder_view 
WHERE Brand = 'AJO'
GROUP BY ID
*/


SELECT A.*, B.ID AS INV_ID, B.Qty AS invoice_total_qty
FROM saOrderListCompareToDelivery_view AS A 
LEFT JOIN (SELECT C.Brand, C.ID, SUM(C.Qty) AS Qty FROM saInv_view AS C GROUP BY C.ID) AS B
ON A.OD_ID = B.ID
-- WHERE A.Brand = 'DYN'
ORDER BY A.RD ASC;

-- SELECT * FROM  saInv_view WHERE ID = '22SS | YWA22C26 | DYN-012 | - | Z1 | BLACK | FREE | PVT | KOREA'

/*
USE sdjo;

DELIMITER //
CREATE PROCEDURE saOrderManagementPeriodSearch_func(IN `str_date` VARCHAR(30), IN `end_date` VARCHAR(30), IN `brand_code` VARCHAR(10))
BEGIN
	IF `brand_code` = '' Then
		SELECT A.*, B.ID AS INV_ID, B.Qty AS invoice_total_qty
		FROM saOrderListCompareToDelivery_view AS A 
		LEFT JOIN (SELECT C.Brand, C.ID, SUM(C.Qty) AS Qty FROM saInv_view AS C GROUP BY C.ID) AS B
		ON A.OD_ID = B.ID
		WHERE A.RD BETWEEN `str_date` AND `end_date`
		ORDER BY A.RD ASC;
	ELSE
		SELECT A.*, B.ID AS INV_ID, B.Qty AS invoice_total_qty
		FROM saOrderListCompareToDelivery_view AS A 
		LEFT JOIN (SELECT C.Brand, C.ID, SUM(C.Qty) AS Qty FROM saInv_view AS C GROUP BY C.ID) AS B
		ON A.OD_ID = B.ID
		WHERE A.RD BETWEEN `str_date` AND `end_date`
		AND A.Brand = `brand_code`
		ORDER BY A.RD ASC;
	END IF;
END;
//
DELIMITER ;
*/

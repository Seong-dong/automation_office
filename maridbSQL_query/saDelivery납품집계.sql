	USE sdjo; 
/*
	SELECT Distinct 
		DATE_FORMAT(a.`Date`,'%m') AS `Date`, 
		a.Brand, 
		sum(a.Amount), 
		Sum(a.Qty) 
	FROM saDelMergeTable AS a 
	WHERE a.Qty Is Not NULL AND
	a.`date` BETWEEN '20220101' AND '20221231'
	GROUP BY a.`date`, a.Brand;
	*/
	
-- DROP PROCEDURE saDelPeriodSearch_func;
/*
DELIMITER //
CREATE PROCEDURE saDelPeriodSearch_func(IN `str_date` VARCHAR(30), IN `end_date` VARCHAR(30), IN `brand_code` VARCHAR(10))
BEGIN
	IF `brand_code` = '' Then
		SELECT Distinct 
			DATE_FORMAT(a.`Date`,'%m') AS `Date`, 
			a.Brand, 
			sum(a.Amount) AS Amount,
			Sum(a.Qty) AS Qty
		FROM saDelMergeTable AS a 
		WHERE a.Qty Is Not Null AND
		a.`date` BETWEEN `str_date` AND `end_date`
		GROUP BY DATE_FORMAT(a.`Date`,'%m'), a.Brand;
	ELSE
		SELECT Distinct 
			DATE_FORMAT(a.`Date`,'%m') AS `Date`, 
			a.Brand, 
			sum(a.Amount) AS Amount, 
			Sum(a.Qty) AS Qty
		FROM saDelMergeTable AS a 
		WHERE a.Qty Is Not Null AND
		a.`date` BETWEEN `str_date` AND `end_date`
		AND Brand = `brand_code`
		GROUP BY DATE_FORMAT(a.`Date`,'%m'), a.Brand;

	END IF;
END;
//
DELIMITER ;
*/

CALL saDelPeriodSearch_func('2022-01-01', '2022-12-31','FLA')
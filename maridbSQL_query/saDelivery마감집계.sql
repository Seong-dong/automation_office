	USE sdjo; 

-- DROP PROCEDURE saInvPeriodSearch_func;

DELIMITER //
CREATE PROCEDURE saInvPeriodSearch_func(IN `str_date` VARCHAR(30), IN `end_date` VARCHAR(30), IN `brand_code` VARCHAR(10))
BEGIN
	IF `brand_code` = '' Then
		SELECT Distinct 
			DATE_FORMAT(a.`Date`,'%m') AS `Date`, 
			a.Brand, 
			sum(a.Amount) AS Amount,
			Sum(a.Qty) AS Qty
		FROM saInvMergeTable AS a 
		WHERE a.Qty Is Not Null AND
		a.`date` BETWEEN `str_date` AND `end_date`
		GROUP BY DATE_FORMAT(a.`Date`,'%m'), a.Brand;
	ELSE
		SELECT Distinct 
			DATE_FORMAT(a.`Date`,'%m') AS `Date`, 
			a.Brand, 
			sum(a.Amount) AS Amount, 
			Sum(a.Qty) AS Qty
		FROM saInvMergeTable AS a 
		WHERE a.Qty Is Not Null AND
		a.`date` BETWEEN `str_date` AND `end_date`
		AND Brand = `brand_code`
		GROUP BY DATE_FORMAT(a.`Date`,'%m'), a.Brand;

	END IF;
END;
//
DELIMITER ;


CALL saInvPeriodSearch_func('2022-01-01', '2022-12-31','FLA')
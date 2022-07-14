USE sdjo;

DROP PROCEDURE saInv_list_func;
DELIMITER //
CREATE PROCEDURE saInv_list_func(IN `str_date` VARCHAR(30), IN `end_date` VARCHAR(30), IN `brand_code` VARCHAR(10))
BEGIN
	IF `brand_code` = '' Then
		SELECT DISTINCT 
			DATE_FORMAT(a.`Date`,'%Y-%m-%d') AS `Date`, 
			a.Season,
			a.`WS No`,
			a.`Style Code`,
			a.`Style Name`,
			a.`Color Code`,
			a.Size,
			a.Factory,
			a.Destination,
			a.Currency,
			a.`Unit Price`,
			a.Amount,
			a.`Total Amount`,
			a.Qty,
			a.Brand 
		FROM saInvMergeTable AS a
		WHERE	a.`Date` BETWEEN `str_date` AND `end_date`
		ORDER BY a.`Date` ASC;
	ELSE
		SELECT DISTINCT 
			DATE_FORMAT(a.`Date`,'%Y-%m-%d') AS `Date`, 
			a.Season,
			a.`WS No`,
			a.`Style Code`,
			a.`Style Name`,
			a.`Color Code`,
			a.Size,
			a.Factory,
			a.Destination,
			a.Currency,
			a.`Unit Price`,
			a.Amount,
			a.`Total Amount`,
			a.Qty,
			a.Brand 
		FROM saInvMergeTable AS a
		WHERE	a.`Date` BETWEEN `str_date` AND `end_date`
		AND a.Brand = `brand_code`
		ORDER BY a.`Date` ASC;
	END IF;
END;
//
DELIMITER ;

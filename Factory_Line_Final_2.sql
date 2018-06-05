

IF EXISTS(SELECT NULL FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Factory_Line_Final') 
DROP TABLE Factory_Line_Final;

----------------------------------------------------------------------------------------
------  Access Crosstab Query   -----------------
----------------------------------------------------------------------------------------

/*

TRANSFORM Sum(Line_Days) AS SumOfLine_Days
SELECT Factory, Line, Sum(Line_Days) AS [Total Of Line_Days]
FROM Factory_Line_temp
GROUP BY Factory, Line
PIVOT DateID;

*/
 

----------------------------------------------------------------------------------------
------ T-SQL Equivalent -----------------
----------------------------------------------------------------------------------------


DECLARE  @ColumnOrder  AS  TABLE(ColumnName  
								varchar(8) 
								NOT NULL 
								PRIMARY KEY) 

DECLARE  @strSQL  AS NVARCHAR(4000)
 
INSERT INTO @ColumnOrder SELECT DISTINCT DateID FROM Factory_Line_temp 


DECLARE  @XTabColumnNames  AS NVARCHAR(MAX)
DECLARE         @XTabColumn      AS varchar(20) 
SET @XTabColumn = (SELECT MIN(ColumnName) FROM  @ColumnOrder) 

SET @XTabColumnNames = N'' 

 
-- Create the xTab columns 

WHILE (@XTabColumn IS NOT NULL) 

  BEGIN 

    SET @XTabColumnNames = @XTabColumnNames + N',' + 
      QUOTENAME(CAST(@XTabColumn AS NVARCHAR(10))) 

   SET @XTabColumn = (SELECT MIN(ColumnName) 
                          FROM   @ColumnOrder 
                          WHERE  ColumnName > @XTabColumn) 
  END 

SET @XTabColumnNames = SUBSTRING(@XTabColumnNames,2,LEN(@XTabColumnNames)) 

PRINT @XTabColumnNames 

SET @strSQL = N'select * into Factory_Line_Final from 
	(SELECT * FROM (
		SELECT Factory, Line, DateID, Line_Days 
		FROM Factory_Line_temp) as header
	pivot (sum(Line_Days)
		FOR DateID IN(' + @XTabColumnNames + N'))  AS Pvt ) as temptable' 

PRINT @strSQL  


-- Execute strSQL sql

EXEC sp_executesql @strSQL 

SELECT * FROM [dbo].[Factory_Line_Final] order by Factory, Line

GO 




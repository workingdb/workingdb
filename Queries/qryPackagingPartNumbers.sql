SELECT PN FROM qryPackagingChildParts
UNION SELECT PN FROM qryActiveProjectPartNumbers GROUP BY PN;


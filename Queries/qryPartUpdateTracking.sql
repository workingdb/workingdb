SELECT tblPartUpdateTracking.recordId, tblPartUpdateTracking.tableName, tblPartUpdateTracking.tableRecordId, tblPartUpdateTracking.updatedBy, tblPartUpdateTracking.updatedDate, tblPartUpdateTracking.columnName, tblPartUpdateTracking.partNumber, tblPartUpdateTracking.dataTag1, tblPartUpdateTracking.dataTag2, grabHistoryRef([previousData],[columnName]) AS previous, grabHistoryRef([newData],[columnName]) AS new
FROM tblPartUpdateTracking
ORDER BY tblPartUpdateTracking.recordId DESC;


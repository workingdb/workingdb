SELECT tblWdbUpdateTracking.recordId, tblWdbUpdateTracking.tableName, tblWdbUpdateTracking.tableRecordId, tblWdbUpdateTracking.updatedBy, tblWdbUpdateTracking.updatedDate, tblWdbUpdateTracking.columnName, grabHistoryRef([previousData],[columnName]) AS previous, grabHistoryRef([newData],[columnName]) AS new, tblWdbUpdateTracking.dataTag0, tblWdbUpdateTracking.dataTag1
FROM tblWdbUpdateTracking
ORDER BY tblWdbUpdateTracking.recordId DESC;


SELECT tblDRSUpdateTracking.tableName, tblDRSUpdateTracking.tableRecordId, tblDRSUpdateTracking.updatedBy, tblDRSUpdateTracking.updatedDate, tblDRSUpdateTracking.columnName, DRShistoryGrabReference([columnName],[previousData]) AS previous, DRShistoryGrabReference([columnName],[newData]) AS new, tblDRSUpdateTracking.dataTag0, tblDRSUpdateTracking.dataTag1
FROM tblDRSUpdateTracking;


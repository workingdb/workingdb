SELECT tbllab_resource_schedule.recordid, tbllab_resource_schedule.resourceid AS id, tbllab_wo_work.partnumber AS A, tblDropDownsSP.lab_work_type AS B, tbllab_resource_schedule.schedulestart AS C, tbllab_resource_schedule.schedulehours, tbllab_resource_schedule.woid
FROM (tbllab_resource_schedule LEFT JOIN tbllab_wo_work ON tbllab_resource_schedule.workid = tbllab_wo_work.recordid) LEFT JOIN tblDropDownsSP ON tbllab_wo_work.worktype = tblDropDownsSP.recordid;


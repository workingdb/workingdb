SELECT tbllab_tech_schedule_alterations.schedulehours, tbllab_tech_schedule_alterations.scheduledate AS C, tbllab_tech_schedule_alterations.username, tbllab_tech_schedule_alterations.scheduletemplateid AS id, tblDropDowns_lab.lab_techschedalt_reason AS A, [schedulehours] & " hrs" AS B
FROM tbllab_tech_schedule_alterations LEFT JOIN tblDropDowns_lab ON tbllab_tech_schedule_alterations.schedulereason = tblDropDowns_lab.recordId;


SELECT tblDropDownsSP.trialresult, tblPartInfo.programId, tblPartInfo.customerId, tblPartInfo.unitId, tblPartTrials.recordId
FROM (tblPartTrials LEFT JOIN tblDropDownsSP ON tblPartTrials.trialResult = tblDropDownsSP.recordid) LEFT JOIN tblPartInfo ON tblPartTrials.partInfoId = tblPartInfo.recordId
WHERE tblPartTrials.recordId > 0;


SELECT tblPrograms.ID AS programId, IIf([gateTitle]="N/A",[partProjectStatus],IIf(Mid([gateTitle],2,1)>=[gateNum],"On Schedule","Behind")) AS ProjectJudgement, qryPartTracker_closedANDopen.projectId
FROM ((tblPartInfo INNER JOIN tblPrograms ON tblPartInfo.programId = tblPrograms.ID) INNER JOIN qryPartTracker_closedANDopen ON tblPartInfo.partNumber = qryPartTracker_closedANDopen.partNumber) INNER JOIN qryProgramReview_expectedGate ON tblPrograms.ID = qryProgramReview_expectedGate.programId;


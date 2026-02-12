SELECT Max(tblProgramEvents.correlatedGate) AS gateNum, tblProgramEvents.programId
FROM tblProgramEvents
WHERE (((tblProgramEvents.eventDate)<Date()))
GROUP BY tblProgramEvents.programId;


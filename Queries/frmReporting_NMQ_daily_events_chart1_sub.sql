SELECT tblProgramEvents.programId, IIf([dataSubmitted],"Submitted","Not Submitted") AS submission, tblProgramEvents.eventDate
FROM tblProgramEvents
WHERE ID > 0;


SELECT tblPartSteps.partNumber, tblPartInfo.description, tblPrograms.modelCode, tblPartInfo.customerPN, tblDropDownsSP.partType, tblPrograms.OEM, tblPrograms.modelName, "N/A" AS gateTitle, "N/A" AS stepType, "N/A" AS Due, tblPartSteps.partProjectId AS projectId, "N/A" AS Notes, sqryPartProjectTags.partProjectStatus, sqryPartProjectTags.partProjectType, tblUnits.unitNumber AS MPunit, tblPartMoldingInfo.assignedPress
FROM (((tblPartSteps LEFT JOIN ((tblPartInfo LEFT JOIN tblPrograms ON tblPartInfo.programId = tblPrograms.ID) LEFT JOIN tblDropDownsSP ON tblPartInfo.partType = tblDropDownsSP.recordid) ON tblPartSteps.partNumber = tblPartInfo.partNumber) INNER JOIN sqryPartProjectTags ON tblPartSteps.partProjectId = sqryPartProjectTags.recordId) LEFT JOIN tblUnits ON tblPartInfo.unitId = tblUnits.recordID) LEFT JOIN tblPartMoldingInfo ON tblPartInfo.moldInfoId = tblPartMoldingInfo.recordId
GROUP BY tblPartSteps.partNumber, tblPartInfo.description, tblPrograms.modelCode, tblPartInfo.customerPN, tblDropDownsSP.partType, tblPrograms.OEM, tblPrograms.modelName, tblPartSteps.partProjectId, "N/A", sqryPartProjectTags.partProjectStatus, sqryPartProjectTags.partProjectType, "N/A", tblUnits.unitNumber, tblPartMoldingInfo.assignedPress
HAVING (
        (
            (Count(tblPartSteps.recordId)) = Count([closeDate])
        )
    );


SELECT findStepDue.partNumber, PI.description, tblPrograms.modelCode, PI.customerPN, DD_PT.partType, tblPrograms.OEM, tblPrograms.modelName, findStepDue.due, PS.stepType, PG.gateTitle, PS.stepDescription AS Notes, PG.projectId, projTags.partProjectStatus, projTags.partProjectType, tblUnits.unitName AS MPunit, PMI.assignedPress
FROM (((SELECT
                    PP.recordId,
                    DD_Status.partProjectStatus,
                    DD_Type.partProjectType
                FROM
                    (
                        (
                            tblPartProject as PP
                            INNER JOIN tblPartProjectTemplate as PT ON PP.projectTemplateId = PT.recordId
                        )
                        INNER JOIN tblDropDownsSP as DD_Status ON PP.projectStatus = DD_Status.recordid
                    )
                    INNER JOIN tblDropDownsSP AS DD_Type ON PT.templateType = DD_Type.recordid
            )  AS projTags INNER JOIN ((tblPartSteps AS PS INNER JOIN tblPartGates AS PG ON PS.partGateId = PG.recordId) INNER JOIN ((((SELECT
                                    PS1.recordId,
                                    PS1.partNumber,
                                    IIf (IsNull (dueDate), gateplanneddate, duedate) AS due,
                                    findStep.MinOfindexOrder,
                                    PS1.indexOrder,
                                    findStep.GateId,
                                    PS1.partGateId
                                FROM
                                    tblPartSteps as PS1
                                    INNER JOIN (
                                        SELECT
                                            min(PS.indexOrder) as MinOfindexOrder,
                                            PS.partNumber,
                                            currentGate.GateId,
                                            currentGate.GatePlannedDate
                                        FROM
                                            tblPartSteps as PS
                                            INNER JOIN (
                                                SELECT
                                                    projectId,
                                                    Min(PG.plannedDate) AS GatePlannedDate,
                                                    Min(PG.recordId) AS GateId
                                                FROM
                                                    tblPartGates as PG
                                                WHERE
                                                    PG.actualDate Is Null
                                                GROUP BY
                                                    projectId
                                            ) as currentGate ON PS.partGateId = currentGate.GateId
                                        WHERE
                                            PS.closeDate Is Null
                                        GROUP BY
                                            PS.partNumber,
                                            GateId,
                                            GatePlannedDate
                                    ) as findStep ON (findStep.GateId = PS1.partGateId)
                                    AND (PS1.indexOrder = findStep.MinOfindexOrder)
                                WHERE
                                    PS1.closeDate Is Null
                            )  AS findStepDue LEFT JOIN tblPartInfo AS PI ON findStepDue.partNumber = PI.partNumber) LEFT JOIN tblDropDownsSP AS DD_PT ON PI.partType = DD_PT.recordid) LEFT JOIN tblPrograms ON PI.programId = tblPrograms.ID) ON PS.recordId = findStepDue.recordId) ON projTags.recordId = PG.projectId) LEFT JOIN tblUnits ON PI.unitId = tblUnits.recordID) LEFT JOIN tblPartMoldingInfo AS PMI ON PI.moldInfoId = PMI.recordId;


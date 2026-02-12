SELECT PS.recordId, PS.partNumber, PS.stepType, PS.responsible, PS.stepDescription, PS.dueDate, tblPartInfo.description, PE.person AS PE, RESP.person AS Responsible_User, tblPartInfo.developingLocation, tblPrograms.modelCode, "G" & [MaxOfcorrelatedGate] AS Correlated_Gate, Left([gateTitle],2) AS currentGate
FROM tblPartGates INNER JOIN (((((tblPartSteps AS PS LEFT JOIN tblPartInfo ON PS.partNumber = tblPartInfo.partNumber) LEFT JOIN (SELECT
                            PT.partNumber,
                            PT.person
                        FROM
                            tblPartTeam as PT
                            INNER JOIN tblPermissions AS P ON PT.person = P.User
                        WHERE
                            (
                                (P.Dept = "Project")
                                AND ((P.Level) = "Engineer")
                            )
                    )  AS PE ON PS.partNumber = PE.partNumber) LEFT JOIN (SELECT
                        PS.recordId,
                        PT.person
                    FROM
                        tblPartSteps as PS
                        INNER JOIN (
                            tblPartTeam AS PT
                            INNER JOIN tblPermissions AS P ON PT.person = P.User
                        ) ON (PS.partNumber = PT.partNumber)
                        AND (PS.responsible = P.Dept)
                    GROUP BY
                        PS.recordId,
                        PT.person,
                        P.Level
                    HAVING
                        (P.Level = "Engineer")
                )  AS RESP ON PS.recordId = RESP.recordId) LEFT JOIN tblPrograms ON tblPartInfo.programId = tblPrograms.ID) LEFT JOIN (SELECT
                programId,
                Max(eventDate) AS MaxOfeventDate,
                Max(correlatedGate) AS MaxOfcorrelatedGate
            FROM
                tblProgramEvents
            GROUP BY
                programId
            HAVING
                Max(eventDate) < Date ()
            ORDER BY
                Max(eventDate)
        )  AS corG ON tblPartInfo.programId = corG.programId) ON tblPartGates.recordId = PS.partGateId
WHERE (((PS.dueDate)<Date()) AND ((PS.closeDate) Is Null)) AND PS.partNumber IN (SELECT partNumber FROM tblPartProject WHERE projectStatus = 1)
ORDER BY PS.recordId DESC;


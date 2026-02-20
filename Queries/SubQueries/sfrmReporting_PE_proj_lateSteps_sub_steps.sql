SELECT tblPartSteps.recordId, tblPartSteps.partNumber, tblPartSteps.responsible, "Complete: " & [stepType] AS [action], tblPartSteps.stepDescription, Nz(tblPartSteps.dueDate,Nz((SELECT
                    TOP 1 s2.dueDate
                FROM
                    tblPartSteps AS s2
                WHERE
                    s2.partGateId = tblPartSteps.partGateId
                    AND (
                        s2.indexOrder > tblPartSteps.indexOrder
                        OR (
                            s2.indexOrder = tblPartSteps.indexOrder
                            AND s2.recordId > tblPartSteps.recordId
                        )
                    )
                    AND s2.dueDate IS NOT NULL
                ORDER BY
                    s2.indexOrder ASC,
                    s2.recordId ASC
            ),tblPartGates.plannedDate)) AS due_date, tblPartInfo.description, PE.person AS PE, RESP.person AS Responsible_User, tblPartInfo.developingLocation, tblPrograms.modelCode
FROM ((SELECT
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
        )  AS PE RIGHT JOIN (((tblPartGates INNER JOIN tblPartSteps ON tblPartGates.recordId = tblPartSteps.partGateId) LEFT JOIN tblPartInfo ON tblPartSteps.partNumber = tblPartInfo.partNumber) LEFT JOIN (SELECT
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
            )  AS RESP ON tblPartSteps.recordId = RESP.recordId) ON PE.partNumber = tblPartSteps.partNumber) LEFT JOIN tblPrograms ON tblPartInfo.programId = tblPrograms.ID
WHERE (((Nz([tblPartSteps].[dueDate],Nz((SELECT
                    TOP 1 s2.dueDate
                FROM
                    tblPartSteps AS s2
                WHERE
                    s2.partGateId = tblPartSteps.partGateId
                    AND (
                        s2.indexOrder > tblPartSteps.indexOrder
                        OR (
                            s2.indexOrder = tblPartSteps.indexOrder
                            AND s2.recordId > tblPartSteps.recordId
                        )
                    )
                    AND s2.dueDate IS NOT NULL
                ORDER BY
                    s2.indexOrder ASC,
                    s2.recordId ASC
            ),[tblPartGates].[plannedDate])))<Date()) AND ((tblPartSteps.status)<>"Closed"));


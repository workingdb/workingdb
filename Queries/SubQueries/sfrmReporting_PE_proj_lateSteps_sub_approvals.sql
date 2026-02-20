SELECT PTA.recordId, PTA.partNumber, PTA.dept & " " & PTA.reqLevel AS responsible, "Approve: " & [stepType] AS [action], tblPartSteps.stepDescription, Nz(tblPartSteps.dueDate,Nz((SELECT
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
FROM (SELECT
            PTA.recordId,
            PT.person
        FROM
            tblPartTrackingApprovals as PTA
            INNER JOIN (
                tblPartTeam AS PT
                INNER JOIN tblPermissions AS P ON PT.person = P.User
            ) ON (PTA.partNumber = PT.partNumber)
            AND (PTA.dept = P.Dept)
            AND (PTA.reqLevel = P.level)
        GROUP BY
            PTA.recordId,
            PT.person,
            P.Level
    )  AS RESP RIGHT JOIN (tblPartTrackingApprovals AS PTA INNER JOIN (((SELECT
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
                )  AS PE RIGHT JOIN ((tblPartGates INNER JOIN tblPartSteps ON tblPartGates.recordId = tblPartSteps.partGateId) LEFT JOIN tblPartInfo ON tblPartSteps.partNumber = tblPartInfo.partNumber) ON PE.partNumber = tblPartSteps.partNumber) LEFT JOIN tblPrograms ON tblPartInfo.programId = tblPrograms.ID) ON PTA.tableRecordId = tblPartSteps.recordId) ON RESP.recordId = PTA.recordId
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


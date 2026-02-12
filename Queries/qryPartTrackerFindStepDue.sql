SELECT PS1.recordId, PS1.partNumber, IIf (IsNull (dueDate), gateplanneddate, duedate) AS due, findStep.MinOfindexOrder, PS1.indexOrder, findStep.GateId, PS1.partGateId
FROM tblPartSteps AS PS1 INNER JOIN (SELECT
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
    )  AS findStep ON (PS1.indexOrder = findStep.MinOfindexOrder) AND (findStep.GateId = PS1.partGateId)
WHERE PS1.closeDate Is Null;


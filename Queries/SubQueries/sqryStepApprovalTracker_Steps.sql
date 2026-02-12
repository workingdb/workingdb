SELECT tblPartSteps.partNumber, tblPartInfo.description, tblPrograms.modelCode, tblDropDownsSP.partType, tblPrograms.OEM, tblPrograms.modelName, tblPartSteps.stepType AS [Action], IIf(
        IsNull([dueDate]),
        IIf(
            IsNull([soonestPillar]),
            IIf(
                IsNull([gateLastPillar]),
                [gatePlannedDate],
                [gateLastPillar]
            ),
            [soonestPillar]
        ),
        [duedate]
    ) AS due, tblPartTeam.person, tblPartSteps.recordId, tblPartGates.gateTitle
FROM qryFindNextPillar INNER JOIN (sqryPartProjectTags INNER JOIN (tblPartGates INNER JOIN (((tblPartSteps LEFT JOIN ((tblDropDownsSP RIGHT JOIN tblPartInfo ON tblDropDownsSP.recordid = tblPartInfo.partType) LEFT JOIN tblPrograms ON tblPartInfo.programId = tblPrograms.ID) ON tblPartSteps.partNumber = tblPartInfo.partNumber) LEFT JOIN tblPartTeam ON tblPartSteps.partNumber = tblPartTeam.partNumber) LEFT JOIN tblPermissions ON tblPartSteps.responsible = tblPermissions.Dept) ON tblPartGates.recordId = tblPartSteps.partGateId) ON sqryPartProjectTags.recordId = tblPartSteps.partProjectId) ON qryFindNextPillar.partGateId = tblPartGates.recordId
WHERE (
        ((tblPartTeam.person) = [User])
        AND ((tblPartSteps.status) <> 'Closed')
        AND ((tblPartSteps.responsible) IS NOT NULL)
        AND (
            (tblPartSteps.indexOrder) < [qryFindNextPillar].[index]
        )
    )
GROUP BY tblPartSteps.partNumber, tblPartInfo.description, tblPrograms.modelCode, tblDropDownsSP.partType, tblPrograms.OEM, tblPrograms.modelName, tblPartSteps.stepType, IIf(
        IsNull([dueDate]),
        IIf(
            IsNull([soonestPillar]),
            IIf(
                IsNull([gateLastPillar]),
                [gatePlannedDate],
                [gateLastPillar]
            ),
            [soonestPillar]
        ),
        [duedate]
    ), tblPartTeam.person, tblPartSteps.recordId, tblPartGates.gateTitle, sqryPartProjectTags.partProjectStatus
HAVING (
        (
            (sqryPartProjectTags.partProjectStatus) <> 'On Hold'
        )
    )
ORDER BY IIf(
        IsNull([dueDate]),
        IIf(
            IsNull([soonestPillar]),
            IIf(
                IsNull([gateLastPillar]),
                [gatePlannedDate],
                [gateLastPillar]
            ),
            [soonestPillar]
        ),
        [duedate]
    );


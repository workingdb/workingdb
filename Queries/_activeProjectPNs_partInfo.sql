SELECT tblPartInfo.*
FROM tblPartInfo INNER JOIN _activeProjectPNs ON tblPartInfo.partNumber = [_activeProjectPNs].partNumber;


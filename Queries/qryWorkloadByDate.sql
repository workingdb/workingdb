SELECT tblWorkloadRanking.userName, tblWorkloadRanking.hoursDate, Round(Sum([hours]),2) AS totalHours
FROM tblWorkloadRanking
GROUP BY tblWorkloadRanking.userName, tblWorkloadRanking.hoursDate
HAVING (((tblWorkloadRanking.hoursDate)>Date()-7));


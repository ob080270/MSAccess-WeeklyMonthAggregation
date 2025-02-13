SELECT 
    Year([smpDate]) AS smpYear, 
    GetMonthForWeek(Year([smpDate]), DatePart("ww", [smpDate], 0)) AS smpMonth, 
    DatePart("ww", [smpDate], 0) AS smpWeek, 

    Sum([smpValue1]) AS smpSumA, 
    Sum([smpValue2]) AS smpSumB

FROM smpTable

GROUP BY 
    Year([smpDate]), 
    GetMonthForWeek(Year([smpDate]), DatePart("ww", [smpDate], 0)), 
    DatePart("ww", [smpDate], 0);

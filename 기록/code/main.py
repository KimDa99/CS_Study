import TimeLineMaker

person = "김다은"
dateRange = ["2024년 03월 1주차", "2024년 04월 1주차"]

doFilterByDate = False
doIncludeMonthlyReview = True
doIncludeWeeklyReview = True

total_sheetName = "전체 한 일"
weeklyCheck_sheetName = "주간 점검"
monthlyCheck_sheetName = "월 점검"
categories_sheetName = "categories"

colorPalette = [
    "FCEBB6","FFD8D8","FFC3A0","FFDAC1","FFB7B2","FFD8CC",
    "FFB6C1","FFC0CB","F8BBD0","E1BEE7","D1C4E9","C5CAE9",
    "BBDEFB","B3E5FC","B2EBF2","B2DFDB","C8E6C9","DCEDC8",
    "F0F4C3","FFF9C4","FFECB3","FFE0B2","FFCCBC","FFAB91"]
dateColorPalette = [

    ["C7D8ED", "F4B3C2", "F2E0B7", "A3D6C1", "E6CF8B"], # January
    ["E7CCE2", "FFD1DC", "F1B9CE", "A9D0F5", "F5E1A8"], # February
    ["C2E0C4", "F9D6E1", "BBD7F1", "FFE8B4", "D8CCEB"], # March
    ["F8E5E5", "AED6F1", "FFF4CF", "C7E2C9", "FAD4A0"], # April
    ["FFDDC1", "C2D6E2", "F3D9E2", "D1E6C9", "F5E1A8"], # May
    ["FFDFBA", "A5D8E2", "F4CFD4", "D1E6C9", "FFD7C1"], # June
    ["FFC2C2", "FFDFBA", "B5D8E2", "FFE8B4", "C7E2C9"], # July
    ["FFD1DC", "BBD7F1", "F3D9E2", "C7E2C9", "FFD7C1"], # August
    ["F3D9E2", "AED6F1", "FFF4CF", "C2E0C4", "FAD4A0"], # September
    ["F5E1A8", "FFC2C2", "D1E6C9", "F4CFD4", "FFDFBA"], # October
    ["C7D8ED", "F4B3C2", "F2E0B7", "A3D6C1", "E6CF8B"], # November
    ["F4B3C2", "C7D8ED", "F2E0B7", "A3D6C1", "E6CF8B"]  # December
]

TimeLineMaker.MakeTimeLine(person=person, dateRange=dateRange, doIncludeMonthlyReview=doIncludeMonthlyReview, 
                           doIncludeWeeklyReview=doIncludeWeeklyReview, doFilterByDate=doFilterByDate, total_sheetName=total_sheetName, 
                           weeklyCheck_sheetName=weeklyCheck_sheetName, monthlyCheck_sheetName=monthlyCheck_sheetName, 
                           categories_sheetName=categories_sheetName, ColorPalette=colorPalette, DateColorPalette=dateColorPalette)
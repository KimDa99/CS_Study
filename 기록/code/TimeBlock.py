class TimeBlock:
    date = "0000년 00월 0주차"

    dataInfoNames = ['날짜', '이름', '분류', 'content', 'Category', 'tag', 'etc']
    data = {'카테고리': {'태그': ['내용1', '내용2']}}

    reviewInfoNames = None #['날짜', '분류', '평가 (잘한 점 + 개선할 점)', '만족도', '기타']
    middleReview = ['날짜', '3/4일 목표', '평가 (잘한 점 + 개선할 점)', '만족도', '기타']
    endReview = ['날짜', '3/4일 목표', '평가 (잘한 점 + 개선할 점)', '만족도', '기타']

    # get input of array of [date, category, tag, content]
    def __init__(self, totalData, reviews = None):
        
        self.dataInfoNames = totalData[0]
        self.date = totalData[1][self.GetColIndex()]

        self.data = {}
        for i in range(1, len(totalData)):
            self.AddData(totalData[i])

        self.middleReview = []
        self.endReview = []
        if reviews is not None:
            self.reviewInfoNames = reviews[0]
            self.AddReviews(reviews)


    def AddDatas(self, datas):
        for data in datas:
            self.AddData(data)

    def AddData(self, data):
        dateCol = self.GetColIndex("날짜")
        date = data[dateCol]
        if date != self.date:
            return False

        CategoryCol = self.GetColIndex("Category")
        category = data[CategoryCol]

        TagCol = self.GetColIndex("tag")
        tag = data[TagCol]

        ContentCol = self.GetColIndex("content")
        content = data[ContentCol]
        
        if category in self.data:
            if tag in self.data[category]:
                self.data[category][tag].append(content)
            else:
                self.data[category][tag] = [content]
        else:
            self.data[category] = {tag: [content]}

    def AddReview(self, review):
        if self.reviewInfoNames is None:
            print("No review info names found")
            return False

        dateCol = self.GetColIndex("날짜", self.reviewInfoNames)
        if review[0] != self.date:
            return False

        divCol = self.GetColIndex("분류", self.reviewInfoNames)

        div = review[divCol]
        if div == "주중":
            self.middleReview = review
        else:
            self.endReview = review

    def AddReviews(self, reviews):
        if self.reviewInfoNames is None:
            self.reviewInfoNames = reviews[0]

        for i in range(1, len(reviews)):
            self.AddReview(reviews[i])

    def GetLength(self):
        maxLen = len(self.review)

        for cat in self.data:
            tmp = 0
            for tag in self.data[cat]:
                tmp += len(self.data[cat][tag])
            if tmp > maxLen:
                maxLen = tmp
        return maxLen

    def GetColIndex(self, name="날짜", arr = dataInfoNames):
        for i in range(len(arr)):
            if arr[i] == name:
                return i
        return -1
    
    def GetTagLength(self, category, tag):
        length = 0
        for k in self.data[category]:
            if k[0] == tag:
                length = len(k[1])

class TimeBlockList:
    TimeBlocks = []

    def __init__(self, datas, reviews=None):
        self.AddDatas(datas)
        if reviews is not None:
            self.AddReviews(reviews)

    def FindBlock(self, date):
        for block in self.TimeBlocks:
            if block.date == date:
                return block
        return None

    def AddData(self, data):
        if(len(data) != 4):
            print("Data format error: ", data)
            return

        block = self.FindBlock(data[0])
        if block is None:
            block = TimeBlock([data])
            self.TimeBlocks.append(block)
        else:
            block.AddData(data)
    
    def AddDatas(self, datas):
        for data in datas:
            self.AddData(data)

    def AddReview(self, review):
        block = self.FindBlock(review[0])
        if block is not None:
            block.AddReview(review)
        else:
            print("No block found for review: ", review)

    def AddReviews(self, reviews):
        for review in reviews:
            self.AddReview(review)

    def sortByDate(self):
        self.TimeBlocks.sort(key=lambda x: x.date)

class MonthBlock:
    month = "0000년 00월"
    TimeBlocks = []
    MiddleMonthReview = []
    LaterMonthReview = []
    Goal = "goal"

    def __init__(self, TimeBlocks, middleMonthReview=[], laterMonthReview=[]):
        self.TimeBlocks = TimeBlocks
        self.month = TimeBlocks[0].date[:8]
        self.AddLaterMonthReview(laterMonthReview)
        self.AddMiddleMonthReview(middleMonthReview)
    
    def AddTimeBlock(self, TimeBlock):
        if TimeBlock.date[:8] == self.month:
            self.TimeBlocks.append(TimeBlock)
        else:
            print("TimeBlock date error: ", TimeBlock.date)
    
    def AddTimeBlocks(self, TimeBlocks):
        for TimeBlock in TimeBlocks:
            self.AddTimeBlock(TimeBlock)
    
    def AddMiddleMonthReview(self, monthReview, dateCol=0):
        if monthReview[dateCol] == self.month:
            self.MiddleMonthReview = monthReview
        else:
            print("Month review date error: ", monthReview)

    def AddLaterMonthReview(self, monthReview, dateCol=0):
        if monthReview[dateCol] == self.month:
            self.LaterMonthReview = monthReview
        else:
            print("Month review date error: ", monthReview)

    def sortByDate(self):
        self.TimeBlocks.sort(key=lambda x: x.date)

class MonthBlockList:
    MonthBlocks = []

    def __init__(self, TimeBlocks, middleMonthReviews=[], laterMonthReviews=[]):
        self.AddTimeBlocks(TimeBlocks)
        self.AddMiddleMonthReviews(middleMonthReviews)
        self.AddLaterMonthReviews(laterMonthReviews)

    def FindBlock(self, month):
        for block in self.MonthBlocks:
            if block.month == month:
                return block
        return None

    def AddTimeBlock(self, TimeBlock):
        month = TimeBlock.date[:8]
        block = self.FindBlock(month)
        if block is None:
            block = MonthBlock([TimeBlock])
            self.MonthBlocks.append(block)
        else:
            block.AddTimeBlock(TimeBlock)

    def AddTimeBlocks(self, TimeBlocks):
        for TimeBlock in TimeBlocks:
            self.AddTimeBlock(TimeBlock)

    def AddMiddleMonthReview(self, monthReview, dateCol=0):
        month = monthReview[0][:8]
        block = self.FindBlock(month)
        if block is not None:
            block.AddMiddleMonthReview(monthReview)
        else:
            print("No block found for review: ", monthReview)

    def AddMiddleMonthReviews(self, monthReviews, dateCol=0):
        for monthReview in monthReviews:
            self.AddMiddleMonthReview(monthReview)

    def AddLaterMonthReview(self, monthReview, dateCol=0):
        month = monthReview[0][:8]
        block = self.FindBlock(month)
        if block is not None:
            block.AddLaterMonthReview(monthReview)
        else:
            print("No block found for review: ", monthReview)

    def AddLaterMonthReviews(self, monthReviews):
        for monthReview in monthReviews:
            self.AddLaterMonthReview(monthReview)
    
    def AddReviews(self, reviews, dateCol=0, divCol=1):
        for review in reviews:
            if review[divCol] == "중간":
                self.AddMiddleMonthReview(review, dateCol)
            elif review[divCol] == "후반":
                self.AddLaterMonthReview(review, dateCol)
            else:
                print("Review division error: ", review)

    def sortByDate(self):
        self.MonthBlocks.sort(key=lambda x: x.month)
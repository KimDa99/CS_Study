class TimeBlock:
    date = "0000년 00월 0주차"
    data = {'카테고리': {'태그': ['내용1', '내용2']}}
    review = {'카테고리': ['성과 및 배운 점', '3/4일 목표', '잘한 점', '개선할 점', '만족도', '기타']}

    # get input of array of [date, category, tag, content]
    def __init__(self, totalData, reviews = None, dateCol = 0, CategoryCol = 1, TagCol = 2, ContentCol = 3):
        dataNum = len(totalData)
        self.date = totalData[0][dateCol]        
        self.data = {}
        self.review = {}
        for i in range(dataNum):
            self.AddData(totalData[i], dateCol, CategoryCol, TagCol, ContentCol)
        if reviews is not None:
            self.AddReviews(reviews)


    def AddDatas(self, datas, dateCol = 0, CategoryCol = 1, TagCol = 2, ContentCol = 3):
        for data in datas:
            self.AddData(data, dateCol, CategoryCol, TagCol, ContentCol)

    def AddData(self, data, dateCol = 0, CategoryCol =1, TagCol = 2, ContentCol = 3):
        date = data[dateCol]
        category = data[CategoryCol]
        tag = data[TagCol]
        content = data[ContentCol]

        if date != self.date:
            return False
        
        if category in self.data:
            if tag in self.data[category]:
                self.data[category][tag].append(content)
            else:
                self.data[category][tag] = [content]
        else:
            self.data[category] = {tag: [content]}


    def AddReview(self, review):
        if review[0] != self.date:
            return False
        category = review[1]
        if category in self.review:
            self.review[category] = review[2:]
        else:
            self.review[category] = review[2:]

    def AddReviews(self, review):
        for r in review:
            self.AddReview(r)

    def GetLength(self):
        maxLen = len(self.review)

        for cat in self.data:
            tmp = 0
            for tag in self.data[cat]:
                tmp += len(self.data[cat][tag])
            if tmp > maxLen:
                maxLen = tmp
        return maxLen

    def GetTagLength(self, category, tag):
        length = 0
        for k in self.data[category]:
            if k[0] == tag:
                length = len(k[1])

class TimeBlockList:
    TimeBlocks = []

    def __init__(self, datas, reviews=None, dateCol = 0, CategoryCol = 1, TagCol = 2, ContentCol = 3):
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
    monthReview = []

    def __init__(self, TimeBlocks, monthReview):
        self.TimeBlocks = TimeBlocks
        self.month = TimeBlocks[0].date[:8]
        if monthReview[0] != self.month:
            print("Month review date error: ", monthReview)
        else:
            self.monthReview = monthReview
    
    def AddTimeBlock(self, TimeBlock):
        if TimeBlock.date[:8] == self.month:
            self.TimeBlocks.append(TimeBlock)
        else:
            print("TimeBlock date error: ", TimeBlock.date)
    
    def AddTimeBlocks(self, TimeBlocks):
        for TimeBlock in TimeBlocks:
            self.AddTimeBlock(TimeBlock)
    
    def AddMonthReview(self, monthReview):
        if monthReview[0] == self.month:
            self.monthReview = monthReview
        else:
            print("Month review date error: ", monthReview)

    def sortByDate(self):
        self.TimeBlocks.sort(key=lambda x: x.date)

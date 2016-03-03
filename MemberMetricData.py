class MemberMetricData():
    # attributes:
    # member_id
    # metric_id
    # metric_data_type
    # metric_value

    def setMemberId(self, member_id):
        self.member_id = member_id

    def getMemberId(self):
        return self.member_id

    def setMetricId(self, metric_id):
        self.metric_id = metric_id

    def getMetricId(self):
        return self.metric_id

    def setMetricDataType(self, metric_data_type):
        self.metric_data_type = metric_data_type

    def getMetricDataType(self):
        return self.metric_data_type

    def setMetricValue(self, metric_value):
        self.metric_value = metric_value

    def getMetricValue(self):
        return self.metric_value

    def print(self):
        print("MemberId[%d], MetricId [%d], MetricDataType[%s], MetricValue[%s]" % (self.getMemberId(), self.getMetricId(), self.getMetricDataType(), self.getMetricValue()))
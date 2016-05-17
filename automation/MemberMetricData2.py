from collections import defaultdict

class MemberMetricData2(object):
    # index to pre-index objects for quick referencing later
    member_metric_index = defaultdict(list)

    # attributes:
    # member_id
    # metric_id
    # metric_data_type
    # metric_value

    def __init__(self, member_id, metric_id, metric_data_type, metric_value):
        self.member_id = member_id
        self.metric_id = metric_id
        self.metric_data_type = metric_data_type
        self.metric_value = metric_value
        MemberMetricData2.member_metric_index[str(member_id) + '_' + str(metric_id)].append(self)

    @classmethod
    def find_by_member_metric_ids(cls, member_id, metric_id):
        return MemberMetricData2.member_metric_index[str(member_id) + '_' + str(metric_id)]

    def getMemberId(self):
        return self.member_id

    def getMetricId(self):
        return self.metric_id

    def getMetricDataType(self):
        return self.metric_data_type

    def getMetricValue(self):
        return self.metric_value

    def print(self):
        try:
            print("MemberId[%d], MetricId[%d], MetricDataType[%s], MetricValue[%s]" % (self.getMemberId(), self.getMetricId(), self.getMetricDataType(), self.getMetricValue()))
        except TypeError as e:
            print("The following values are being skipped: ", self.getMemberId(), self.getMetricId(), self.getMetricDataType(), self.getMetricValue())
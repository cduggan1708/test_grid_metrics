class MetricSelectorMapper():
    # attributes:
    # metric_id
    # metric_display_name
    # metric_program_name
    # metric_top_level_selector

    def __init__(self, metric_id, metric_display_name, metric_program_name, metric_top_level_selector):
        self.metric_id = metric_id
        self.metric_display_name = metric_display_name
        self.metric_program_name = metric_program_name
        self.metric_top_level_selector = metric_top_level_selector

    def getMetricId(self):
        return self.metric_id

    def getMetricDisplayName(self):
        return self.metric_display_name

    def getMetricProgramName(self):
        return self.metric_program_name

    def getMetricTopLevelSelector(self):
        return self.metric_top_level_selector

    def print(self):
        try:
            print("MetricId [%s], MetricDisplayName[%s], MetricProgramName[%s], MetricTopLevelSelector[%s]" % (self.getMetricId(), self.getMetricDisplayName(), self.getMetricProgramName(), self.getMetricTopLevelSelector()))
        except TypeError as e:
            print("The following values are being skipped: ", self.getMetricId(), self.getMetricDisplayName(), self.getMetricProgramName(), self.getMetricTopLevelSelector())
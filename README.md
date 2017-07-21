# TrendsCalculationDemo
Core methods for calculating trends:
  DataAnalysis.GetTheilRegressionValues()
  DataAnalysis.GetKendallTheilTrendStatistics()
  
Driver method:
  TrendSummaryTable.ascx.vb::BuildTrendTableMarkup()
  
Test page:
  TrendsTable.aspx
  
Primary sequence of core operations:
  1. TrendsTable.aspx instantiates a TrendSummaryTable user control
  2. The TrendSummaryTable user control retrieves the relevant base data from a database
  3. The TrendSummaryTable user control invokes DataAnalysis.GetTheilRegressionValues()
  4. DataAnalysis.GetTheilRegressionValues() invokes DataAnalysis.GetKendallTheilTrendStatistics()
  5. The TrendSummaryTable user control renders the resulting trend results to the TrendsTable.aspx page

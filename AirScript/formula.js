结案日期 = IF(OR(ISBLANK([问题分流日期]), ISBLANK([结案日期]), NOT(OR([状态] = "已完成", [状态] = "已转交"))), "", NETWORKDAYS([问题分流日期], [结案日期]))
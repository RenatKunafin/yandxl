from yametrics import Yametrics
from excel import Excel

ym = Yametrics()
data = ym.request_metrics()

e = Excel(data)
e.init_wb()
# e.write_to_wb()

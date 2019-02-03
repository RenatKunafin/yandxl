from yametrics import Yametrics

# metrics = [
#     # 'ym:s:visits',
#     # 'ym:s:users',
#     # 'ym:s:pageviews',
#     'ym:s:goal<goal_id>reaches',
#     'ym:s:goal<goal_id>visits',
#     'ym:s:goal<goal_id>users'
# ]

ym = Yametrics()
ym.request_metrics()
# ym.get_logs()

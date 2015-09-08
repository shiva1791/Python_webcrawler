__author__ = 'Shiva'
from time import localtime

activities = {8: 'Sleeping',
              9: 'Commuting',
              17: 'Working',
              18: 'Commuting',
              20: 'Eating',
              24: 'Resting' }

time_now = localtime()
print(time_now.tm_hour,':',time_now.tm_min,':',time_now.tm_sec)
hour = time_now.tm_hour

for activity_time in sorted(activities.keys()):
    if hour < activity_time:
        print(activities[activity_time])
        break
else:
    print('Unknown, AFK or sleeping!')
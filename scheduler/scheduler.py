from apscheduler.schedulers.blocking import BlockingScheduler

sched=BlockingScheduler()
sched.add_job(parse,'interval',minutes=5)

try:
    sched.start()
except:
    print "Could not run script"
    sys.exit()

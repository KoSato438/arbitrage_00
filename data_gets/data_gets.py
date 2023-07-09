import main
import schedule,time,datetime
import os

try:
    os.makedirs("./datas/GMO")
    os.makedirs("./datas/coincheck")
    os.makedirs("./datas/liquid")
    os.makedirs("./datas/all")
except FileExistsError:
    pass

def job():
    now=datetime.datetime.now()
    main.main()

schedule.every(1).seconds.do(job)

while True:
    schedule.run_pending()
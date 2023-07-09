import main
import schedule,time,datetime
import os

new_dir_path_recursive="./transactions"
try:
    os.makedirs(new_dir_path_recursive)
except FileExistsError:
    pass

def job():
    now=datetime.datetime.now()
    main.main()
#    try:
#        check.main()
#    except:
#        print("何かしらのエラーでmain()失敗")
#    print("<<DONE>>",now.strftime('%Y-%m-%d %H:%M:%S'))

schedule.every(5).seconds.do(job)

while True:
    schedule.run_pending()
    time.sleep(1)

check.main()
from alive_progress import alive_bar
import time

mylist = [1,2,3,4,5,6,7,8]

with alive_bar(len(mylist)) as bar:
    for i in mylist:
        bar()
        time.sleep(1)
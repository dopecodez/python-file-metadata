import win32com.client
import itertools
import threading
import time
import sys
from operator import itemgetter

sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0)
ns = sh.NameSpace(r'E:\Music\Alternative(Indie too!)')
colnum = 0
columns = []
while True:
    colname=ns.GetDetailsOf(None, colnum)
    if not colname:
        break
    columns.append(colname)
    colnum += 1

done = False
#here is the animation
def animate():
    for c in itertools.cycle(['|', '/', '-', '\\']):
        if done:
            break
        sys.stdout.write('\rloading ' + c)
        sys.stdout.flush()
        time.sleep(0.1)
    sys.stdout.write('\rDone!     ')

t = threading.Thread(target=animate)
t.start()

authors = []
for item in ns.Items():
    for colnum in range(len(columns)):
        colval=ns.GetDetailsOf(item, colnum)
        if colval:
            if(columns[colnum] == "Authors"):
                names= colval.split(',')
                for name in names:
                    authors.append(name)

#long process here
time.sleep(10)
done = True

countedOrders = [[x,authors.count(x)] for x in set(authors)]
countedOrders = sorted(countedOrders, key=itemgetter(1), reverse=True)

for item in countedOrders:
    output = '\n' + 'Artist:' + item[0] + " | Number Of Tracks:" + str(item[1]) + '\n'
    print(output)

            
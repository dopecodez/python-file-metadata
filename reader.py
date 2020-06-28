import win32com.client
import itertools
import threading
import time
import sys
from operator import itemgetter

sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0)
ns = sh.NameSpace(r'E:\Music\Rap & Hip-Hop')
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

t = threading.Thread(target=animate)
t.start()

authors = []
for item in ns.Items():
    for colnum in range(len(columns)):
        colval=ns.GetDetailsOf(item, colnum)
        if colval:
            if(columns[colnum] == "Authors"):
                colval = colval.replace(";",",").replace("&",",").replace("feat.", ",")
                names = colval.split(',')
                for name in names:
                    authors.append(name.rstrip().lstrip())

#long process here
time.sleep(10)
done = True

countedOrders = [[x,authors.count(x)] for x in set(authors)]
countedOrders = sorted(countedOrders, key=itemgetter(1), reverse=True)

for item in countedOrders:
    output = '\n\n' + item[0] + " | " + str(item[1]) + '\n'
    print(output)

            
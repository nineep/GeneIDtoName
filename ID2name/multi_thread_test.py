import threading
import time

d = {1: 'a', 2: 'b', 3: 'c', 4: 'd', 5: 'c', 6: 'e', 7: 'f', 8: 'g'}


def popit(d,i):
    print()
    print("Execution of Thread {} started".format(i))
    print('pop前字典', d)
    if d:
        d.popitem()
    else:
        print('字典空了')
    print('pop后字典', d)
    time.sleep(3)
    print("Execution of Thread {} finished".format(i))


while d:
    for i in range(4):
        if len(d) > 0:
            print('字典长度:', len(d))
            thread = threading.Thread(target=popit, args=(d, i,))
            thread.start()
            #print("Active Threads:", threading.enumerate())
time.sleep(3)
print('最终:', d)
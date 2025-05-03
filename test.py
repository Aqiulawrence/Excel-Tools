import threading

shared_list = []
def unsafe_append():
    for _ in range(1000):
        shared_list.append(1)

threads = [threading.Thread(target=unsafe_append) for _ in range(1000)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(len(shared_list))  # 可能小于 10000（数据丢失！）
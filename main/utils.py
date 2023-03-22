import time
from queue import Empty


def countdown(seconds: int) -> None:
    while seconds >= 0:
        m, s = divmod(seconds, 60)
        timer = f'{m:02d}:{s:02d}'
        print('Time until file gets deleted:', timer, end='\r')
        time.sleep(1)
        seconds -= 1


def execute_queue(queue) -> None:
    while True:
        try:
            q = queue.get()
            q()
        except Empty:
            pass
"""防止电脑进入锁屏：定时轻微移动鼠标（移动后立刻复位）。

用法:
    python keep_awake.py                     # 默认每 30 秒微动一次
    python keep_awake.py -i 10               # 每 10 秒微动一次
    python keep_awake.py --duration 3600     # 运行 1 小时后自动退出
按 Ctrl+C 停止。
"""

import argparse
import ctypes
import random
import time
from datetime import datetime

MOUSEEVENTF_MOVE = 0x0001

user32 = ctypes.WinDLL("user32", use_last_error=True)
user32.mouse_event.argtypes = [
    ctypes.c_uint,      # dwFlags
    ctypes.c_long,      # dx
    ctypes.c_long,      # dy
    ctypes.c_uint,      # dwData
    ctypes.c_size_t,    # dwExtraInfo
]


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


def get_cursor_pos():
    p = POINT()
    user32.GetCursorPos(ctypes.byref(p))
    return p.x, p.y


def jiggle():
    """向右或左移动 1~2 像素，再原路移回（净位移为 0）。"""
    dx = random.choice((-1, 1)) * random.randint(1, 2)
    user32.mouse_event(MOUSEEVENTF_MOVE, dx, 0, 0, 0)
    time.sleep(0.05)
    user32.mouse_event(MOUSEEVENTF_MOVE, -dx, 0, 0, 0)


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def main():
    parser = argparse.ArgumentParser(description="定时微动鼠标，防止电脑锁屏")
    parser.add_argument("-i", "--interval", type=int, default=30,
                        help="微动间隔秒数，默认 30")
    parser.add_argument("-d", "--duration", type=int, default=0,
                        help="运行多少秒后自动退出，默认 0 表示一直运行")
    args = parser.parse_args()

    if args.interval < 1:
        parser.error("间隔至少要 1 秒")

    last_pos = get_cursor_pos()
    start = time.monotonic()
    count = 0
    print(f"[{now()}] 已启动：每 {args.interval} 秒微动鼠标一次，Ctrl+C 退出。",
          flush=True)

    try:
        while True:
            time.sleep(args.interval)

            cur = get_cursor_pos()
            if cur != last_pos:
                # 鼠标位置变了，说明本人在操作电脑，无需干预
                last_pos = cur
                continue

            jiggle()
            count += 1
            print(f"[{now()}] 第 {count} 次微动鼠标，位置 {cur}", flush=True)

            if args.duration and time.monotonic() - start >= args.duration:
                print(f"[{now()}] 已运行 {args.duration} 秒，自动退出。", flush=True)
                break
    except KeyboardInterrupt:
        print(f"\n[{now()}] 已停止，共微动 {count} 次。", flush=True)


if __name__ == "__main__":
    main()

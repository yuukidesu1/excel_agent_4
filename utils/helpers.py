import time
import functools

def timer(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()

        # 执行原函数
        result = func(*args, **kwargs)

        end_time = time.perf_counter()
        duration = end_time - start_time
        print(f"DEBUG: [函数 {func.__name__}] 运行耗时 {duration:.4f} 秒")

        return result
    return wrapper
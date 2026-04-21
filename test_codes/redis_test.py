import redis

# 测试默认配置
r1 = redis.Redis(host='localhost', port=6379, db=0)
try:
    print("测试 localhost:6379 db=0:", r1.ping())
except Exception as e:
    print(f"localhost连接失败: {e}")

# 测试其他可能配置
r2 = redis.Redis(host='127.0.0.1', port=6379, db=0)
try:
    print("测试 127.0.0.1:6379 db=0:", r2.ping())
except Exception as e:
    print(f"127.0.0.1连接失败: {e}")

# 测试是否需要密码
r3 = redis.Redis(host='localhost', port=6379, db=0, password=None)
try:
    print("测试无密码:", r3.ping())
except Exception as e:
    print(f"无密码连接失败: {e}")


value = 0x000000  # 分割したい値（例: 24ビット値）

# 上位8ビットを取得
upper = (value >> 16) & 0xFF

# 中位8ビットを取得
middle = (value >> 8) & 0xFF

# 下位8ビットを取得
lower = value & 0xFF

# 結果を出力
print(f"Upper: {upper:#04x}")
print(f"Middle: {middle:#04x}")
print(f"Lower: {lower:#04x}")

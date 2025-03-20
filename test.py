import pandas as pd

# 读取原始Excel文件
input_file = "故障工单2025-03-20 13_32_32.xlsx"
output_file = "output.xlsx"

df = pd.read_excel(input_file)

# 处理房间号字段（假设原始列名为"房间号"）
def extract_room(room_str):
    try:
        parts = room_str.split('_')
        return '_'.join(parts[:2])  # 提取前两部分
    except:
        return None

df['房间'] = df['房间'].apply(extract_room)

def parse_datetime(dt_str):
    # 如果值是 <nil> 或空值，返回 NaT（Not a Time）
    if pd.isna(dt_str) or dt_str == "<nil>":
        return pd.NaT
    try:
        # 尝试解析 ISO 8601 格式时间
        return pd.to_datetime(dt_str)
    except:
        return pd.NaT

# 解析维修开始时间列
df['维修开始时间'] = df['维修开始时间'].apply(parse_datetime)

# 拆分维修开始时间为日期和时间
df['维修日期'] = df['维修开始时间'].dt.date
df['维修时间'] = df['维修开始时间'].dt.time

# 创建新DataFrame并按指定顺序排列列
new_df = pd.DataFrame({
    '工单ID': df['ID'],
    '所属物理机': df['资产编号'],
    '库房位置': df['房间'],
    '服务器厂商': df['厂商'],
    '服务器SN': df['SN'],
    '故障盘品牌': df['原件品牌'],
    '故障盘SN': df['原件SN'],
    '故障盘PN': df['原件PN'],
    '更换盘品牌': df['新件品牌'],
    '更换盘SN': df['新件SN'],
    '更换盘PN': df['新件PN'],
    '维修日期': df['维修日期'],
    '维修时间': df['维修时间'],
    '不返还字段':'不返还',
    '故障盘配件容量':'',
    '故障盘型号':'',
    'IP': df['IPv6'],
    '机柜位置': df['机架位']
})

# 保存到新Excel文件
new_df.to_excel(output_file, index=False, engine='openpyxl')

print(f"处理完成，已保存到 {output_file}")
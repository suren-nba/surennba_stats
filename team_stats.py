import os
import time
import pandas as pd
import requests

# ==========================================================================
# 1. 基础设置与目录准备 (开源版相对路径配置)
# ==========================================================================
# 自动在当前运行目录下创建 data/team_stats 文件夹进行分类存储
output_dir = os.path.join(os.getcwd(), "data", "team_stats")

if not os.path.exists(output_dir):
    os.makedirs(output_dir, exist_ok=True)

# ==========================================================================
# 2. 定义抓取任务队列 (常规赛 + 季后赛全量生成，Excel 格式)
# ==========================================================================
# 命名逻辑: 2025-26NBA_(SeasonType)_(Team/Opponent)_stats.xlsx
api_tasks = [
    # --- 常规赛 (Regular Season) ---
    {
        "type_name": "常规赛-球队本队",
        "url": "https://api.pbpstats.com/get-totals/nba?Season=2025-26&SeasonType=Regular%2BSeason&StartType=All&Type=Team",
        "save_path": os.path.join(
            output_dir, "2025-26NBA_RegularSeason_Team_stats.xlsx"
        ),
    },
    {
        "type_name": "常规赛-球队对手",
        "url": "https://api.pbpstats.com/get-totals/nba?Season=2025-26&SeasonType=Regular%2BSeason&StartType=All&Type=Opponent",
        "save_path": os.path.join(
            output_dir, "2025-26NBA_RegularSeason_Opponent_stats.xlsx"
        ),
    },
    # --- 季后赛 (Playoffs) ---
    {
        "type_name": "季后赛-球队本队",
        "url": "https://api.pbpstats.com/get-totals/nba?Season=2025-26&SeasonType=Playoffs&StartType=All&Type=Team",
        "save_path": os.path.join(
            output_dir, "2025-26NBA_Playoffs_Team_stats.xlsx"
        ),
    },
    {
        "type_name": "季后赛-球队对手",
        "url": "https://api.pbpstats.com/get-totals/nba?Season=2025-26&SeasonType=Playoffs&StartType=All&Type=Opponent",
        "save_path": os.path.join(
            output_dir, "2025-26NBA_Playoffs_Opponent_stats.xlsx"
        ),
    },
]

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "application/json",
    "Origin": "https://www.pbpstats.com",
    "Referer": "https://www.pbpstats.com/",
}

# ==========================================================================
# 3. 核心抓取循环
# ==========================================================================
print("开始同步 2025-26 赛季全量数据...")
print(f"存储目录: {output_dir}\n" + "=" * 50)

for task in api_tasks:
    print(f"正在抓取: {task['type_name']} ...")

    try:
        response = requests.get(task["url"], headers=headers, timeout=60)

        if response.status_code == 200:
            json_data = response.json()
            table_data = json_data.get("multi_row_table_data", [])

            if len(table_data) > 0:
                df = pd.DataFrame(table_data)
                df.columns = df.columns.str.lower()

                # 同时检测旧版的 object 类型和新版的纯 string 类型，确保正确对齐缩进
                for col in df.select_dtypes(
                    include=["object", "string"]
                ).columns:
                    df[col] = df[col].astype(str)

                # 核心保存逻辑：输出为 Excel 格式
                df.to_excel(task["save_path"], index=False)

                print(f" 成功！获取 {len(df)} 行数据。")
                print(f" 文件名: {os.path.basename(task['save_path'])}")
            else:
                print(" 解析结果为空，该赛段可能暂无数据。")
        else:
            print(f"   请求失败：HTTP {response.status_code}")

        # 礼貌性延时：非常重要，防止高频访问被 pbpstats 封禁 IP
        time.sleep(2)
        print("-" * 50)

    except Exception as e:
        print(f" 发生异常: {e}\n" + "-" * 50)

print("\n任务完成！")
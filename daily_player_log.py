import os
import pandas as pd
from nba_api.stats.endpoints import leaguegamelog

# ==============================================================================
# 第一部分：网络环境配置 
# ==============================================================================
# 提示：因国内网络环境特殊，若抓取时遇到 Timeout 报错，请在此处填入你的科学上网代理地址
# 示例：os.environ['http_proxy'] = 'http://127.0.0.1:7890'
os.environ['http_proxy'] = ''
os.environ['https_proxy'] = ''
os.environ['HTTP_PROXY'] = ''
os.environ['HTTPS_PROXY'] = ''

print("Status: Starting high-speed data ingestion (EXCEL mode)...")

try:
    # ==============================================================================
    # 第二部分：调用 NBA 官方 API 抓取数据 (常规赛、附加赛、季后赛)
    # ==============================================================================
    
    # 定义我们需要抓取的所有赛季类型
    season_types = ["Regular Season", "PlayIn", "Playoffs"]
    all_frames = []
    
    for st in season_types:
        print(f"正在尝试获取: {st} 数据...")
        try:
            # 传入 season_type_all_star 参数
            log = leaguegamelog.LeagueGameLog(
                season="2025-26", 
                season_type_all_star=st,  
                player_or_team_abbreviation="P",
                timeout=100
            )
            
            df_st = log.get_data_frames()[0]
            
            if not df_st.empty:
                df_st['SEASON_TYPE'] = st 
                all_frames.append(df_st)
                print(f"  -> 抓取到 {len(df_st)} 条 {st} 记录")
            else:
                print(f"  ->  {st} 目前无数据")
                
        except Exception as inner_e:
            print(f"  -> 获取 {st} 时出现异常 (通常是因为该阶段尚未开始)，跳过。")

    if len(all_frames) == 0:
        raise ValueError("所有赛季类型均未抓取到数据，请检查网络！")

    # 将三个阶段的数据无缝合并成一张大表
    df = pd.concat(all_frames, ignore_index=True)
    
    # ==============================================================================
    # 第三部分：数据清洗与时区转换
    # ==============================================================================
    
    # 1. 字段名全小写，方便后续分析
    df.columns = df.columns.str.lower()
    
    # 2. 处理日期：将美国东部时间 (ET) 转换为国内时间 (+1天)
    df['game_date'] = pd.to_datetime(df['game_date'])
    df['game_date'] = df['game_date'] + pd.Timedelta(days=1)
    
    # 3. 确保最新日期的比赛永远在第一行
    df = df.sort_values(by="game_date", ascending=False)
    
    # 4. 转换回标准的年-月-日字符串
    df['game_date'] = df['game_date'].dt.strftime('%Y-%m-%d')
    
    # ==============================================================================
    # 第四部分：本地持久化保存 
    # ==============================================================================
    
    # 核心修改：使用相对路径，自动在脚本所在目录下创建一个 "data" 文件夹
    target_dir = os.path.join(os.getcwd(), "data")
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
        
    # 保存为 .csv 格式
    file_path = os.path.join(target_dir, "nba_daily_live.csv")
    
    # 输出到 CSV 
    df.to_csv(file_path, index=False, encoding='utf_8_sig')
    
    print("-" * 50)
    print("Success: Data synchronization completed successfully.")
    print(f"Total Rows Ingested: {len(df)}")
    print(f"Latest Game Date: {df['game_date'].iloc[0]}")
    print(f"Format: Comma-Separated Values (.csv)")
    print(f"Destination: {file_path}")
    print("-" * 50)
except Exception as e:
    print("-" * 50)
    print(f"Error: Data fetch failed.")
    print(f"Reason: {e}")
    print("-" * 50)

input("\nProcess finished. Press Enter to close this window...")
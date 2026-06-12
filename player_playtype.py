import os
import sys
import time
import pandas as pd
from nba_api.stats.endpoints import synergyplaytypes

# ==============================================================================
# 第一部分：环境配置与跨平台相对路径 (Excel 格式)
# ==============================================================================
# 自动在当前运行目录下寻找或创建 data 文件夹
OUTPUT_DIR = os.path.join(os.getcwd(), "data")
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

FILE_NAME = "2025-26_player_playtype.xlsx"
FILE_PATH = os.path.join(OUTPUT_DIR, FILE_NAME)

# 临时取消代理设置，防止由于网络环境导致 API 请求被拦截
os.environ['http_proxy'] = ''
os.environ['https_proxy'] = ''
os.environ['HTTP_PROXY'] = ''
os.environ['HTTPS_PROXY'] = ''

print("Status: Starting Synergy [PLAYER] PlayType data ingestion...")
print(f"Target Path: {FILE_PATH}")

# ==============================================================================
# 第二部分：智能增量更新逻辑 (基于“是否有季后赛数据”判断)
# ==============================================================================
CURRENT_SEASON = "2025-26" 
TARGET_SEASONS = ['Regular Season', 'Playoffs']
seasons_to_fetch = TARGET_SEASONS 
existing_rs_df = pd.DataFrame()

if os.path.exists(FILE_PATH):
    try:
        # 变更为读取 Excel
        temp_df = pd.read_excel(FILE_PATH)
        
        # 1. 统一将列名转为小写进行条件判断
        temp_df.columns = temp_df.columns.str.lower()
        
        # 2. 增强容错：清理列内容中可能存在的前后空格，并转换为全小写
        has_playoffs_by_text = False
        has_playoffs_by_id = False
        
        if 'season_type_custom' in temp_df.columns:
            has_playoffs_by_text = temp_df['season_type_custom'].astype(str).str.strip().str.lower().eq('playoffs').any()
            
        if 'season_id' in temp_df.columns:
            has_playoffs_by_id = temp_df['season_id'].astype(str).str.strip().eq('42025').any()
        
        # 核心判断：只要满足文本或 ID 任意一个条件，即代表本地已存在季后赛数据
        if has_playoffs_by_text or has_playoffs_by_id:
            print("--> 检测到已存在【季后赛】球员数据")
            print("--> 仅更新季后赛，保留常规赛历史数据。")
            
            # 提取并永久保留常规赛部分
            is_rs_mask = temp_df['season_type_custom'].astype(str).str.strip().str.lower() == 'regular season'
            existing_rs_df = temp_df[is_rs_mask].copy()
            
            # 任务列表：只抓季后赛
            seasons_to_fetch = ['Playoffs']
        else:
            print("--> 本地尚未发现季后赛球员数据，将进行双轨全量扫描...")
            seasons_to_fetch = TARGET_SEASONS
            
    except Exception as e:
        print(f"--> 本地文件读取异常: {e}，将尝试全量抓取...")
else:
    print("--> 本地无历史文件，准备初始化抓取全量球员数据...")

PLAY_TYPES = [
    'Isolation', 'Transition', 'PRBallHandler', 'PRRollman', 
    'Postup', 'Spotup', 'Handoff', 'Cut', 'OffScreen', 'OffRebound', 'Misc'
]
TYPE_GROUPINGS = ['offensive', 'defensive']
new_data_frames = []

# ==============================================================================
# 第三部分：核心抓取循环
# ==============================================================================
try:
    for s_type in seasons_to_fetch:
        for grouping in TYPE_GROUPINGS:
            for p_type in PLAY_TYPES:
                print(f"Fetching: [{s_type}] [{grouping.upper()}] - {p_type} ...", end=" ", flush=True)
                
                try:
                    synergy_data = synergyplaytypes.SynergyPlayTypes(
                        league_id='00',
                        per_mode_simple='PerGame',
                        player_or_team_abbreviation='P', # P 代表球员级数据
                        season_type_all_star=s_type,  
                        season=CURRENT_SEASON,
                        play_type_nullable=p_type,
                        type_grouping_nullable=grouping,
                        timeout=100
                    )
                    
                    df = synergy_data.get_data_frames()[0]
                    
                    if not df.empty:
                        # 强制将新抓取数据的列名转为小写，确保融合时完美对齐
                        df.columns = df.columns.str.lower()
                        
                        df['play_type_custom'] = p_type
                        df['type_grouping_custom'] = grouping
                        df['season_type_custom'] = s_type 
                        new_data_frames.append(df)
                        print(f"成功 ({len(df)} 球员)")
                    else:
                        print("为空")
                
                except Exception as inner_e:
                    print(f"失败 (跳过)。原因: {inner_e}")
                
                # 防封锁请求间隔
                time.sleep(2) 
                
    # ==============================================================================
    # 第四部分：数据融合与持久化保存 (输出为 Excel)
    # ==============================================================================
    if new_data_frames:
        recent_fetch_df = pd.concat(new_data_frames, ignore_index=True)
        
        # 融合逻辑：保留的常规赛 + 最新抓取的数据
        if not existing_rs_df.empty:
            final_df = pd.concat([existing_rs_df, recent_fetch_df], ignore_index=True)
        else:
            final_df = recent_fetch_df
            
        # 核心去重主键：以 player_id 为核心基准，结合招式维度去重
        final_df = final_df.drop_duplicates(subset=['player_id', 'play_type_custom', 'type_grouping_custom', 'season_type_custom'])
        
        # Pandas 3.0 兼容处理：将所有文本/对象列强转为标准字符串，防止 Excel 引擎报错
        for col in final_df.select_dtypes(include=["object", "string"]).columns:
            final_df[col] = final_df[col].astype(str)
            
        # 写入本地 Excel 归档
        final_df.to_excel(FILE_PATH, index=False)
        
        print("-" * 50)
        print("任务完成！")
        print(f"数据库总行数: {len(final_df)}")
        print(f"数据已保存至: {FILE_PATH}")
    else:
        print("警告: 未获取到任何新数据。")

except Exception as e:
    print(f"核心流程发生严重错误: {e}")
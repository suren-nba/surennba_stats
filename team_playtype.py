import os
import time
import pandas as pd
from nba_api.stats.endpoints import synergyplaytypes

# ==============================================================================
# 1. 环境配置与文件路径
# ==============================================================================
# 自动在当前运行目录下创建 data/team_stats 文件夹
output_dir = os.path.join(os.getcwd(), "data", "team_stats")
if not os.path.exists(output_dir):
    os.makedirs(output_dir, exist_ok=True)

# 设定输出文件名 (改为 Excel 格式)
file_name = "2025-26NBA_Synergy_stats.xlsx"
file_path = os.path.join(output_dir, file_name)

# 国内网络环境如果出现超时，请在此处填写科学上网代理
os.environ['http_proxy'] = ''
os.environ['https_proxy'] = ''
os.environ['HTTP_PROXY'] = ''
os.environ['HTTPS_PROXY'] = ''

print("状态: 开始抓取 Synergy 进攻/防守方式数据 (Excel 模式)...")

# ==============================================================================
# 2. 智能增量更新逻辑 (基于本地文件状态判断)
# ==============================================================================
current_season = "2025-26"
target_seasons = ['Regular Season', 'Playoffs']
seasons_to_fetch = target_seasons 
existing_rs_df = pd.DataFrame()

# 如果本地已经存在历史数据文件，尝试读取以决定抓取策略
if os.path.exists(file_path):
    try:
        temp_df = pd.read_excel(file_path)
        
        # 核心判断逻辑：如果本地数据中已包含季后赛标签，说明当前处于季后赛阶段
        if 'Playoffs' in temp_df['season_type_custom'].values:
            print("--> 检测到本地已存在 [季后赛] 数据。")
            print("--> 增量更新策略生效：仅抓取季后赛数据，保留历史常规赛数据。")
            
            # 提取并缓存常规赛部分，等待后续合并
            existing_rs_df = temp_df[temp_df['season_type_custom'] == 'Regular Season'].copy()
            # 缩减抓取任务目标
            seasons_to_fetch = ['Playoffs']
        else:
            print("--> 本地仅包含常规赛数据，将进行全量扫描 (常规赛 + 季后赛)...")
            seasons_to_fetch = target_seasons
            
    except Exception as e:
        print(f"--> 本地文件读取异常: {e}。策略回退：执行全量抓取...")
else:
    print("--> 首次运行，未检测到本地数据。准备执行全量抓取...")

# Synergy 定义的 11 种核心回合类型
play_types = [
    'Isolation', 'Transition', 'PRBallHandler', 'PRRollman', 
    'Postup', 'Spotup', 'Handoff', 'Cut', 'OffScreen', 'OffRebound', 'Misc'
]
# 分别抓取进攻端与防守端
type_groupings = ['offensive', 'defensive']
new_data_frames = []

# ==============================================================================
# 3. 核心抓取循环
# ==============================================================================
try:
    for s_type in seasons_to_fetch:
        for grouping in type_groupings:
            for p_type in play_types:
                print(f"抓取中: [{s_type}] [{grouping.upper()}] - {p_type} ...", end=" ", flush=True)
                
                try:
                    synergy_data = synergyplaytypes.SynergyPlayTypes(
                        league_id='00',
                        per_mode_simple='PerGame',
                        player_or_team_abbreviation='T',
                        season_type_all_star=s_type,  
                        season=current_season,
                        play_type_nullable=p_type,
                        type_grouping_nullable=grouping,
                        timeout=100
                    )
                    
                    df = synergy_data.get_data_frames()[0]
                    
                    if not df.empty:
                        # 补充自定义标签，方便后续数据分析时进行筛选
                        df['play_type_custom'] = p_type
                        df['type_grouping_custom'] = grouping
                        df['season_type_custom'] = s_type 
                        new_data_frames.append(df)
                        print(f"成功 ({len(df)} 行)")
                    else:
                        print("空数据")
                
                except Exception as inner_e:
                    print(f"抓取失败并跳过。原因: {inner_e}")
                
                # 礼貌性延时，避免触发 NBA 官网的反爬虫风控
                time.sleep(2) 
                
    # ==============================================================================
    # 4. 数据融合与持久化保存
    # ==============================================================================
    if new_data_frames:
        recent_fetch_df = pd.concat(new_data_frames, ignore_index=True)
        recent_fetch_df.columns = recent_fetch_df.columns.str.lower()
        
        # 融合逻辑：保留的常规赛历史数据 + 刚刚抓取的新数据
        if not existing_rs_df.empty:
            final_df = pd.concat([existing_rs_df, recent_fetch_df], ignore_index=True)
        else:
            final_df = recent_fetch_df
            
        # 根据核心业务主键进行最终去重，确保数据纯净
        final_df = final_df.drop_duplicates(
            subset=['team_id', 'play_type_custom', 'type_grouping_custom', 'season_type_custom']
        )
        
        # 兼容性处理：防止老版本 Pandas 的 object 类型导致 Excel 写入失败
        for col in final_df.select_dtypes(include=["object", "string"]).columns:
            final_df[col] = final_df[col].astype(str)
        
        # 写入本地 Excel 文件
        final_df.to_excel(file_path, index=False)
        
        print("-" * 50)
        print("任务完成！")
        print(f"数据库总行数: {len(final_df)}")
        print(f"文件存储位置: {file_path}")
    else:
        print("警告: 本次运行未能抓取到任何新数据。")

except Exception as e:
    print(f"发生严重错误: {e}")

input("\n按回车键关闭窗口...")
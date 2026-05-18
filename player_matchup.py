import os
import pandas as pd
import time
from nba_api.stats.endpoints import leagueseasonmatchups
from nba_api.stats.endpoints import boxscorematchupsv3
from nba_api.stats.static import teams

# ==============================================================================
# 1. 基础设置与相对路径
# ==============================================================================
# 自动在当前运行目录下创建 data/matchups 文件夹
output_dir = os.path.join(os.getcwd(), "data", "matchups")
if not os.path.exists(output_dir):
    os.makedirs(output_dir, exist_ok=True)

reg_file_path = os.path.join(output_dir, "2025-26NBA_Regular_Matchups.xlsx")
po_file_path = os.path.join(output_dir, "2025-26NBA_Playoffs_Matchups.xlsx")
# 之前生成的每日单场球员数据，也就是说如果要生成季后赛对位必须先更新球员每日单场数据
daily_live_path = os.path.join(os.getcwd(), "data", "nba_daily_live.xlsx")

API_SEASON = "2025-26" 
SLEEP_TIME = 1.5       

os.environ['http_proxy'] = ''
os.environ['https_proxy'] = ''
os.environ['HTTP_PROXY'] = ''
os.environ['HTTPS_PROXY'] = ''

# ==============================================================================
# 2. 核心抓取模块 (常规赛)
# ==============================================================================
def fetch_regular_season():
    """遍历 30 支球队，抓取常规赛整体对位数据"""
    nba_teams = teams.get_teams()
    team_ids = [team['id'] for team in nba_teams]
    all_matchups = []
    
    print("\n抓取模块: [常规赛 整体汇总]")
    for i, team_id in enumerate(team_ids):
        print(f"[{i+1}/30] 请求球队ID: {team_id}...", end=" ", flush=True)
        try:
            matchup_log = leagueseasonmatchups.LeagueSeasonMatchups(
                season=API_SEASON, 
                season_type_playoffs="Regular Season", 
                per_mode_simple="Totals", 
                def_team_id_nullable=team_id, 
                timeout=100
            )
            df_current = matchup_log.get_data_frames()[0]
            
            if not df_current.empty:
                all_matchups.append(df_current)
                print("成功")
            else:
                print("为空")
            
            time.sleep(SLEEP_TIME) 
            
        except Exception as e:
            print(f"失败 (跳过)。原因: {e}")
            continue

    if all_matchups:
        res_df = pd.concat(all_matchups, ignore_index=True)
        res_df.columns = res_df.columns.str.lower()
        res_df.drop_duplicates(inplace=True)
        return res_df
    return pd.DataFrame()

# ==============================================================================
# 3. 核心抓取模块 (季后赛 - 单场 GameID 嗅探模式)
# ==============================================================================
def fetch_playoffs_by_game():
    """通过读取 nba_daily_live.xlsx 获取季后赛 Game ID，逐场抓取"""
    print("\n抓取模块: [季后赛 单场探测]")
    
    if not os.path.exists(daily_live_path):
        print("提示: 找不到本地 nba_daily_live.xlsx，无法获取季后赛单场 ID。")
        return pd.DataFrame()

    try:
        df_live = pd.read_excel(daily_live_path)
        df_live.columns = df_live.columns.str.lower()
        
        # 提取属于季后赛的唯一 Game ID
        if 'season_type' in df_live.columns and 'game_id' in df_live.columns:
            playoff_games = df_live[df_live['season_type'] == 'Playoffs']['game_id'].unique()
        else:
            return pd.DataFrame()
            
    except Exception as e:
        print(f"读取 daily_live 文件失败: {e}")
        return pd.DataFrame()

    if len(playoff_games) == 0:
        print("提示: 雷达文件内无季后赛记录，说明季后赛尚未开始，安全跳过。")
        return pd.DataFrame()

    print(f"-->检测到 {len(playoff_games)} 场季后赛，开始逐场请求...")
    
    all_po_matchups = []
    for i, gid in enumerate(playoff_games):
        # 补齐 10 位数，防止 Excel 把 '0042200101' 吞掉前置零变成整数
        gid_str = str(gid).zfill(10)
        print(f"[{i+1}/{len(playoff_games)}] 请求 GameID: {gid_str}...", end=" ", flush=True)
        
        try:
            # 调用单场对位接口
            matchup = boxscorematchupsv3.BoxScoreMatchupsV3(game_id=gid_str, timeout=100)
            df_game = matchup.get_data_frames()[0]
            
            if not df_game.empty:
                df_game['source_game_id'] = gid_str # 打下单场烙印
                all_po_matchups.append(df_game)
                print("成功")
            else:
                print("为空")
                
            time.sleep(SLEEP_TIME)
            
        except Exception as e:
            print(f"失败 ({e})")

    if all_po_matchups:
        res_df = pd.concat(all_po_matchups, ignore_index=True)
        res_df.columns = res_df.columns.str.lower()
        return res_df
    return pd.DataFrame()

# ==============================================================================
# 4. 智能增量状态机与持久化保存
# ==============================================================================
print("状态: 启动对位数据嗅探器...")

need_fetch_regular = True

if os.path.exists(po_file_path):
    print("--> 本地已存在 [季后赛] 数据。")
    print("--> 判定处于季后赛时段：将停止常规赛抓取，仅更新季后赛场次。")
    need_fetch_regular = False
else:
    print("--> 本地无季后赛对位数据。准备双轨并进...")

try:
    # --- 分支 A：常规赛 ---
    if need_fetch_regular:
        df_reg = fetch_regular_season()
        if not df_reg.empty:
            # Pandas 3.0 Excel 兼容处理
            for col in df_reg.select_dtypes(include=["object", "string"]).columns:
                df_reg[col] = df_reg[col].astype(str)
            df_reg.to_excel(reg_file_path, index=False)
            print(f"[{reg_file_path}] 常规赛固化完成。行数: {len(df_reg)}")

    # --- 分支 B：季后赛 ---
    df_po = fetch_playoffs_by_game()
    if not df_po.empty:
        # Pandas 3.0 Excel 兼容处理
        for col in df_po.select_dtypes(include=["object", "string"]).columns:
            df_po[col] = df_po[col].astype(str)
        df_po.to_excel(po_file_path, index=False)
        print(f"[{po_file_path}] 季后赛固化完成。行数: {len(df_po)}")
    else:
        print("--> 季后赛输出拦截：未获取到有效数据，不生成空文件。")

    print("\n" + "=" * 50)
    print("对位数据同步结束！")
    print("=" * 50)

except Exception as e:
    print(f"\n发生严重运行时错误: {e}")

input("\n按回车键关闭窗口...")
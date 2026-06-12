import os
import pandas as pd
import time
import sys
from nba_api.stats.endpoints import boxscorematchupsv3
from nba_api.stats.endpoints import leagueseasonmatchups
from nba_api.stats.static import teams

# ==============================================================================
# 1. 基础设置与跨平台相对路径配置
# ==============================================================================
output_dir = os.path.join(os.getcwd(), "data", "matchups")
if not os.path.exists(output_dir):
    os.makedirs(output_dir, exist_ok=True)

# 雷达文件：在 data 根目录下
daily_live_path = os.path.join(os.getcwd(), "data", "nba_daily_live.xlsx")

# 产出与缓存文件：在 data/matchups 目录下
reg_file_path = os.path.join(output_dir, "2025-26NBA_Regular_Matchups.xlsx")
output_file_path = os.path.join(output_dir, "2025-26NBA_Playoffs_Matchups.xlsx")
cache_file_path = os.path.join(output_dir, "matchup_playoffs_raw_cache.xlsx")

API_SEASON = "2025-26"
SLEEP_TIME = 1.0

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
    
    print("\nStatus: 开始抓取常规赛整体汇总数据...")
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
# 3. (感知季后赛，并决定是否跳过常规赛)
# ==============================================================================
print("Status: 启动对位数据中控引擎...")

need_fetch_regular = True

if os.path.exists(output_file_path):
    print(f"--> 本地已存在季后赛数据 ({os.path.basename(output_file_path)})。")
    print("--> 判定当前处于季后赛时段：将冻结常规赛耗时抓取，仅执行季后赛增量。")
    need_fetch_regular = False
else:
    print("--> 本地无季后赛对位数据。准备双轨并进...")

if need_fetch_regular:
    df_reg = fetch_regular_season()
    if not df_reg.empty:
        # Pandas 3.0 兼容：防止 Excel 写入引擎崩溃
        for col in df_reg.select_dtypes(include=["object", "string"]).columns:
            df_reg[col] = df_reg[col].astype(str)
            
        df_reg.to_excel(reg_file_path, index=False)
        print(f"[{os.path.basename(reg_file_path)}] 常规赛固化完成。行数: {len(df_reg)}")
    else:
        print("常规赛数据抓取为空。")

# ==============================================================================
# 4. 季后赛增量抓取 (读取赛程 -> 查缓存 -> 抓新增 -> 聚合)
# ==============================================================================
print("\nStatus: 启动季后赛增量抓取引擎...")
try:
    if not os.path.exists(daily_live_path):
        print(f"Error: 找不到基础赛程文件: {os.path.basename(daily_live_path)}，已安全退出。")
        sys.exit(0)
        
    # 核心坑位防御：强制要求 pandas 把 game_id 作为字符串读取，防止吞掉前置 0
    df_live = pd.read_excel(daily_live_path, dtype={'game_id': str})
    df_live.columns = df_live.columns.str.lower()
    
    # 提取季后赛标识
    if 'season_type' in df_live.columns:
        df_playoffs = df_live[df_live['season_type'] == 'Playoffs'].copy()
    else:
        # 兜底：如果没找到 season_type，靠 game_id 的首位数字 4 判断
        df_playoffs = df_live[df_live['game_id'].astype(str).str.startswith('004')].copy()
    
    if df_playoffs.empty:
        print("Status: 雷达表中尚未发现季后赛比赛记录，终止季后赛抓取。")
    else:
        all_game_ids = df_playoffs['game_id'].astype(str).unique().tolist()
        
        # 读取缓存以确定已抓取场次
        processed_ids = []
        if os.path.exists(cache_file_path):
            df_cache = pd.read_excel(cache_file_path, dtype={'gameId': str})
            processed_ids = df_cache['gameId'].astype(str).unique().tolist()
            
        target_ids = [gid for gid in all_game_ids if gid not in processed_ids]
        print(f"Status: 总计 {len(all_game_ids)} 场，已抓取 {len(processed_ids)} 场，待抓取 {len(target_ids)} 场。")

        # 遍历拉取新增数据
        new_matchups = []
        if target_ids:
            for i, gid in enumerate(target_ids):
                # 补齐 10 位确保万无一失
                gid_str = str(gid).zfill(10)
                print(f"[{i+1}/{len(target_ids)}] Fetching GameID: {gid_str}...", end=" ", flush=True)
                try:
                    matchup_log = boxscorematchupsv3.BoxScoreMatchupsV3(game_id=gid_str, timeout=100)
                    df_current = matchup_log.get_data_frames()[0]
                    if not df_current.empty:
                        df_current['gameId'] = gid_str # 用于后续缓存记录
                        new_matchups.append(df_current)
                        print("成功")
                    else:
                        print("为空")
                    time.sleep(SLEEP_TIME)
                except Exception as e:
                    print(f"失败。原因: {e}")

        # 合并缓存、清洗与聚合
        # 1. 更新原始缓存
        if new_matchups:
            df_new = pd.concat(new_matchups, ignore_index=True)
            if os.path.exists(cache_file_path):
                df_old = pd.read_excel(cache_file_path, dtype={'gameId': str})
                df_all = pd.concat([df_old, df_new], ignore_index=True)
            else:
                df_all = df_new
                
            # Pandas 3.0 Excel 兼容
            for col in df_all.select_dtypes(include=["object", "string"]).columns:
                df_all[col] = df_all[col].astype(str)
            df_all.to_excel(cache_file_path, index=False)
        else:
            if os.path.exists(cache_file_path):
                df_all = pd.read_excel(cache_file_path, dtype={'gameId': str})
            else:
                print("没有新数据且无缓存，流程结束。")
                df_all = pd.DataFrame()

        if not df_all.empty:
            # 2. 清洗与聚合 
            df_raw = df_all.copy()
            df_raw['firstNameOff'] = df_raw['firstNameOff'].fillna('')
            df_raw['familyNameOff'] = df_raw['familyNameOff'].fillna('')
            df_raw['firstNameDef'] = df_raw['firstNameDef'].fillna('')
            df_raw['familyNameDef'] = df_raw['familyNameDef'].fillna('')
            
            df_raw['off_player_name'] = (df_raw['firstNameOff'] + ' ' + df_raw['familyNameOff']).str.strip()
            df_raw['def_player_name'] = (df_raw['firstNameDef'] + ' ' + df_raw['familyNameDef']).str.strip()
            
            rename_dict = {
                'personIdOff': 'off_player_id', 'personIdDef': 'def_player_id',
                'partialPossessions': 'partial_poss', 'playerPoints': 'player_pts',
                'teamPoints': 'team_pts', 'matchupAssists': 'matchup_ast',
                'matchupTurnovers': 'matchup_tov', 'matchupBlocks': 'matchup_blk',
                'matchupFieldGoalsMade': 'matchup_fgm', 'matchupFieldGoalsAttempted': 'matchup_fga',
                'matchupThreePointersMade': 'matchup_fg3m', 'matchupThreePointersAttempted': 'matchup_fg3a',
                'helpBlocks': 'help_blk', 'helpFieldGoalsMade': 'help_fgm',
                'helpFieldGoalsAttempted': 'help_fga', 'matchupFreeThrowsMade': 'matchup_ftm',
                'matchupFreeThrowsAttempted': 'matchup_fta', 'shootingFouls': 'sfl'
            }
            
            df_renamed = df_raw.rename(columns=rename_dict)
            numeric_cols = [
                'partial_poss', 'player_pts', 'team_pts', 'matchup_ast', 
                'matchup_tov', 'matchup_blk', 'matchup_fgm', 'matchup_fga', 
                'matchup_fg3m', 'matchup_fg3a', 'help_blk', 'help_fgm', 
                'help_fga', 'matchup_ftm', 'matchup_fta', 'sfl'
            ]
            
            for col in numeric_cols:
                if col in df_renamed.columns:
                    df_renamed[col] = pd.to_numeric(df_renamed[col], errors='coerce').fillna(0)
                    
            df_agg = df_renamed.groupby(
                ['off_player_id', 'off_player_name', 'def_player_id', 'def_player_name'], 
                as_index=False
            ).agg(
                gp=('gameId', 'count'), 
                **{col: (col, 'sum') for col in numeric_cols if col in df_renamed.columns}
            )
            
            df_agg['season_id'] = '42025'
            
            # Pandas 3.0 Excel 兼容
            for col in df_agg.select_dtypes(include=["object", "string"]).columns:
                df_agg[col] = df_agg[col].astype(str)
                
            df_agg.to_excel(output_file_path, index=False)
            
            print(f"Success! 聚合对位数据已更新至 {os.path.basename(output_file_path)}")

except Exception as e:
    print(f"Error: {e}")
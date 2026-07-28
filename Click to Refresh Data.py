import re
import openpyxl
import json

try:
    import requests
except ImportError:
    requests = None

EXCEL_FILE = 'INPUT.xlsx'
JS_FILE = 'enemies.js'


def safe_float(value, default=0.0):
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return default


def safe_int(value, default=0):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def load_existing_data():
    """从现有 enemies.js 读取数据，返回 dict {name: player}"""
    try:
        with open(JS_FILE, 'r', encoding='utf-8') as f:
            js_text = f.read()
        # enemies.js 内容: var playersData = [...];
        json_str = js_text[js_text.index('['):js_text.rindex(']')+1]
        # 兼容 JS 对象字面量语法：裸属性名加引号，移除尾随逗号
        json_str = re.sub(r'(?<=[{,])(\s*)(\w+)(?=\s*:)', r'\1"\2"', json_str)
        json_str = re.sub(r',(\s*)(?=[}\]])', r'\1', json_str)
        arr = json.loads(json_str)
        return {p['name']: p for p in arr}
    except (FileNotFoundError, json.JSONDecodeError, ValueError):
        return {}


def read_excel():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    except FileNotFoundError:
        print(f"错误：找不到文件 {EXCEL_FILE}")
        return []
    except Exception as e:
        print(f"读取Excel时出错: {e}")
        return []

    ws = wb.active
    if ws is None:
        print("Excel文件中没有工作表")
        return []

    players = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 25:
            continue
        (name, number, level, dmg_type, hp, atk, def_, mres, interval, range_,
         speed, dph, dps, burn_dps, phys_resist, magic_resist, dmg_resist,
         silence, stun, frozen, disarmed_combat, palsy, cost, ability,
         description) = row[:25]

        player = {
            "name": str(name) if name else "",
            "icon": "",
            "number": str(number) if number is not None else "",
            "level": str(level) if level else "",
            "dmgType": str(dmg_type) if dmg_type else "",
            "hp": safe_int(hp),
            "atk": safe_float(atk),
            "def": safe_float(def_),
            "mres": safe_float(mres),
            "interval": safe_float(interval),
            "range": safe_float(range_),
            "speed": safe_float(speed),
            "dph": safe_float(dph),
            "dps": safe_float(dps),
            "burnDps": safe_float(burn_dps),
            "physResist": safe_float(phys_resist),
            "magicResist": safe_float(magic_resist),
            "dmgResist": safe_float(dmg_resist),
            "silence": str(silence) if silence else "无",
            "stun": str(stun) if stun else "无",
            "frozen": str(frozen) if frozen else "无",
            "disarmedCombat": str(disarmed_combat) if disarmed_combat else "无",
            "palsy": str(palsy) if palsy else "无",
            "cost": str(cost) if cost is not None else "",
            "ability": str(ability).replace('\n', '<br>') if ability else "",
            "description": str(description) if description else ""
        }
        players.append(player)
    return players


def get_icon_from_api(name):
    if not requests:
        return None
    filename = f"头像_敌人_{name}.png"
    try:
        resp = requests.get(
            "https://prts.wiki/api.php",
            params={
                "action": "query",
                "titles": f"File:{filename}",
                "prop": "imageinfo",
                "iiprop": "url",
                "format": "json"
            },
            timeout=10
        )
        resp.raise_for_status()
        data = resp.json()
        pages = data.get("query", {}).get("pages", {})
        for page_info in pages.values():
            if "missing" not in page_info:
                imageinfo = page_info.get("imageinfo", [])
                if imageinfo:
                    return imageinfo[0]["url"]
    except requests.Timeout:
        print(f"  API请求超时：{name}")
    except requests.RequestException as e:
        print(f"  API请求失败：{name} - {e}")
    except json.JSONDecodeError as e:
        print(f"  API返回数据异常：{name} - {e}")
    return None


def get_number_index_map():
    """从 PRTS 敌人一览/数据页获取编号映射 {name: enemyIndex}"""
    if not requests:
        return {}
    try:
        resp = requests.get(
            "https://prts.wiki/index.php?title=敌人一览/数据&action=raw",
            timeout=15
        )
        resp.raise_for_status()
        raw_data = resp.json()
        return {item['name']: item['enemyIndex'] for item in raw_data}
    except requests.Timeout:
        print("  警告：连接PRTS超时，无法获取编号数据")
        return {}
    except requests.RequestException as e:
        print(f"  警告：获取PRTS编号数据失败 - {e}")
        return {}
    except json.JSONDecodeError:
        print("  警告：PRTS编号数据格式异常")
        return {}


def main():
    print("读取 Excel 数据...")
    players = read_excel()
    if not players:
        print("没有读取到任何数据")
        return

    # 从现有 enemies.js 读取旧数据（头像和编号）
    old_data = load_existing_data()

    # 从 PRTS 数据页获取编号映射 {name: enemyIndex}，作为第三层回退
    number_index_map = get_number_index_map()
    if number_index_map:
        print(f"  从PRTS获取编号数据：{len(number_index_map)} 条")

    missing_icons = []
    missing_numbers = []

    for p in players:
        name = p['name']

        # 头像 - 优先从旧 enemies.js 获取，其次尝试 API
        if name in old_data and old_data[name].get('icon'):
            p['icon'] = old_data[name]['icon']
        elif requests:
            icon = get_icon_from_api(name)
            if icon:
                p['icon'] = icon
                print(f"  API获取头像成功：{name}")
            else:
                p['icon'] = ""
                missing_icons.append(name)
        else:
            p['icon'] = ""
            missing_icons.append(name)

        # 编号 - 优先级：旧JS > Excel > PRTS数据页 > 手动补充
        # p['number'] 已通过 read_excel 初始化为 Excel 值
        js_number = old_data[name].get('number', '') if name in old_data else ''

        if js_number and not p['number']:
            # 旧JS有而Excel无 → 用旧JS值
            p['number'] = js_number
        elif js_number and p['number'] and js_number != p['number']:
            # 冲突 → Excel覆盖，打印提示
            print(f"  编号冲突，使用Excel值：{name} ({js_number} -> {p['number']})")
        # 其余情况：保留Excel值（已在第70行初始化）

        if not p['number']:
            # 第三层：从PRTS数据页查询
            api_num = number_index_map.get(name, '')
            if api_num:
                p['number'] = api_num
                print(f"  从PRTS数据获取编号成功：{name} -> {api_num}")
            else:
                missing_numbers.append(name)

    print("写入 enemies.js...（按Excel行序）")
    with open(JS_FILE, 'w', encoding='utf-8') as f:
        f.write('var playersData = ')
        json.dump(players, f, ensure_ascii=False, indent=2)
        f.write(';\n')

    print("\n完成！")
    if missing_icons:
        print("以下单位的头像未找到（已留空），请手动补充：")
        for n in missing_icons:
            print(f"  - {n}")
    if missing_numbers:
        print("以下单位的编号未找到（已留空）：")
        for n in missing_numbers:
            print(f"  - {n}")
    if not missing_icons and not missing_numbers:
        print("所有头像和编号皆已就位。")


if __name__ == "__main__":
    main()
    input("所有任务已完成，按回车键退出程序...")

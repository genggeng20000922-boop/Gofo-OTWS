# -*- coding: utf-8 -*-
"""
供应商维度数据分析报表生成脚本 (0608-0614)
每个供应商对应不同需求仓的发单数量及发单占比
"""

import re
from collections import defaultdict

# ============== 截图数据（发单需求数 + 线上派遣数）— 原始数据 ==============
RAW_DEMAND_DATA = [
    ("CNO.H", "JC-SL Express Inc", 706, 699),
    ("CNO.H", "J&S Out Front delivery and Transportation LLC-15%", 245, 244),
    ("CNO.H", "Select Hire Services LLC", 56, 34),
    ("CNO.H", "Direct Jobs Solutions Inc", 270, 228),
    ("CNO.H", "AAS Express LLC-11.76%", 84, 62),
    ("CNO.H", "A-Share Express LLC-6%", 119, 111),
    ("CNO.H", "JC-SL Express Inc-15%", 1, 0),
    ("LAV.H", "TG Express TRUCKER LLC-5.25%", 346, 222),
    ("LAV.H", "JC-SL Express Inc", 317, 204),
    ("LAV.H", "Select Hire Services LLC", 21, 12),
    ("LAV.H", "Reliant HR LLC-17%", 118, 78),
    ("DEN", "April Flower LLC-20%", 376, 247),
    ("PHX", "Select Hire Services LLC", 1, 0),
    ("SFO.H", "Flash Way Express LLC-10%", 554, 493),
    ("SEA.H", "Select Hire Services LLC", 249, 249),
    ("SEA.H", "Imperial Workforce LLC", 184, 184),
]


def strip_percent_suffix(name):
    """去掉 -XX% 后缀，前面一致则视为同一供应商"""
    return re.sub(r'-\d+(\.\d+)?%$', '', name).strip()


def build_merged_demand():
    """合并同一仓库+去后缀供应商的需求数和线上派遣数"""
    merged = defaultdict(lambda: [0, 0])
    for wh, sp, demand, online in RAW_DEMAND_DATA:
        sp_clean = strip_percent_suffix(sp)
        key = (wh, sp_clean)
        merged[key][0] += demand
        merged[key][1] += online
    return {k: tuple(v) for k, v in merged.items()}


DEMAND_DATA = build_merged_demand()


def build_data():
    """按需求仓维度聚合，每个仓下列出供应商"""
    by_warehouse = defaultdict(lambda: {
        'suppliers': [],  # list of (supplier, demand, online)
        'total_demand': 0,
        'total_online': 0,
    })
    # Also collect supplier set for card summary
    all_suppliers = set()

    for (wh, sp), (demand, online) in DEMAND_DATA.items():
        by_warehouse[wh]['suppliers'].append((sp, demand, online))
        by_warehouse[wh]['total_demand'] += demand
        by_warehouse[wh]['total_online'] += online
        all_suppliers.add(sp)

    # Add offline dispatch data
    OFFLINE_TARGET_IDS = [
        'GL501046', 'GL500158', 'GL501432', 'GL502878', 'GL501153',
        'GL001219', 'GL501564', 'GL000890', 'GL502030'
    ]

    from openpyxl import load_workbook

    def normalize_supplier(name):
        name = name.strip()
        name = re.sub(r'^\(WE\)\s*', '', name)
        name = strip_percent_suffix(name)
        name = name.replace('Epxress', 'Express')
        return name

    def match_warehouse(temp_site, wh):
        site_upper = temp_site.upper()
        if wh == "CNO.H": return "CNO" in site_upper
        elif wh == "LAV.H": return "LAV" in site_upper and "LAV-HUB" in site_upper
        elif wh == "DEN": return "DEN" in site_upper and ".H" in site_upper
        elif wh == "PHX": return "PHX" in site_upper
        elif wh == "SFO.H": return "SFO" in site_upper
        elif wh == "SEA.H": return "SEA" in site_upper
        elif wh == "SAN01": return "SAN01" in site_upper
        return False

    wb = load_workbook(r'D:/AI/文件/临时工派遣 (61).xlsx')
    ws = wb.active
    offline_by_supplier_wh = defaultdict(int)

    for row in range(2, ws.max_row + 1):
        creator_id = ws.cell(row=row, column=21).value
        if creator_id and str(creator_id).strip() in OFFLINE_TARGET_IDS:
            site = str(ws.cell(row=row, column=18).value or '')
            supplier = normalize_supplier(str(ws.cell(row=row, column=8).value or ''))
            found = False
            for (wh, sp) in DEMAND_DATA:
                if sp == supplier and match_warehouse(site, wh):
                    offline_by_supplier_wh[(sp, wh)] += 1
                    found = True
                    break
            # Check for SAN01 unmatched records
            if not found and match_warehouse(site, "SAN01"):
                offline_by_supplier_wh[(supplier, "SAN01")] += 1

    # Add SAN01 as a warehouse if it has offline records
    san01_suppliers = []
    for (sp, wh), offline in offline_by_supplier_wh.items():
        if wh == "SAN01" and offline > 0:
            san01_suppliers.append((sp, 0, 0, offline))
    if san01_suppliers:
        by_warehouse["SAN01"] = {
            'suppliers': san01_suppliers,
            'total_demand': 0,
            'total_online': 0,
            'total_offline': sum(s[3] for s in san01_suppliers)
        }
        for sp in set(s[0] for s in san01_suppliers):
            all_suppliers.add(sp)

    # Add offline data and sort suppliers within each warehouse
    total_offline_all = 0
    for wh, data in by_warehouse.items():
        if wh == "SAN01":
            total_offline_all += data['total_offline']
            continue
        offline_total = 0
        for i, (sp, demand, online) in enumerate(data['suppliers']):
            offline = offline_by_supplier_wh.get((sp, wh), 0)
            data['suppliers'][i] = (sp, demand, online, offline)
            offline_total += offline
        data['total_offline'] = offline_total
        total_offline_all += offline_total
        # Sort suppliers within warehouse by 线上派遣率 descending
        data['suppliers'].sort(key=lambda x: x[2] / (x[2] + x[3]) if (x[2] + x[3]) > 0 else -1, reverse=True)

    return by_warehouse, all_suppliers, total_offline_all


def generate_html(by_warehouse, all_suppliers, total_offline):
    """生成需求仓维度HTML报表"""
    # Sort warehouses: HUBs first by order, then sites
    hub_order = ["CNO.H", "LAV.H", "SFO.H", "SEA.H"]
    site_order = ["DEN", "PHX", "SAN01"]

    def wh_sort_key(wh):
        if wh in hub_order:
            return (0, hub_order.index(wh))
        elif wh in site_order:
            return (1, site_order.index(wh))
        else:
            return (2, wh)

    sorted_warehouses = sorted(by_warehouse.items(), key=lambda x: wh_sort_key(x[0]))

    total_demand = sum(d['total_demand'] for d in by_warehouse.values())
    total_online = sum(d['total_online'] for d in by_warehouse.values())

    total_fulfill_rate = total_online / total_demand * 100 if total_demand > 0 else 0
    total_online_rate = total_online / (total_online + total_offline) * 100 if (total_online + total_offline) > 0 else 0

    hubs = [w for w in hub_order if w in by_warehouse]
    sites = [w for w in site_order if w in by_warehouse]

    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>美西OTWS门户推广数据分析报表-0608-0614</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: #f5f7fa;
            padding: 20px;
            color: #333;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        h1 {{
            text-align: center;
            color: #1a3a5c;
            margin-bottom: 30px;
            font-size: 28px;
        }}
        .summary {{
            margin-bottom: 30px;
        }}
        .rate-row {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 16px;
            margin-bottom: 4px;
        }}
        .number-row {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 16px;
            margin-bottom: 4px;
        }}
        .area-row {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 16px;
        }}
        .section-label {{
            grid-column: 1 / -1;
            font-size: 14px;
            font-weight: 600;
            color: #999;
            margin-bottom: -4px;
            margin-top: 12px;
        }}
        .card {{
            background: white;
            border-radius: 10px;
            padding: 18px 14px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            border: 1px solid #f0f0f0;
            text-align: center;
        }}
        .card-value {{
            font-size: 28px;
            font-weight: bold;
            color: #1a3a5c;
            margin: 6px 0;
        }}
        .card-label {{
            color: #999;
            font-size: 13px;
        }}

        table {{
            width: 100%;
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            border-collapse: collapse;
        }}
        th {{
            background: linear-gradient(135deg, #1a3a5c 0%, #2c5282 100%);
            color: white;
            padding: 16px 12px;
            font-weight: 600;
            font-size: 14px;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        td {{
            padding: 14px 12px;
            border-bottom: 1px solid #f0f0f0;
            font-size: 14px;
        }}
        tr:hover {{ background: #f8fafc; }}

        .num {{ text-align: right; font-family: 'SF Mono', 'Consolas', monospace; }}
        .rate {{ text-align: center; }}

        .badge {{
            display: inline-block;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 500;
        }}
        .badge-success {{ background: #e6f7ed; color: #1a9c4a; }}
        .badge-warning {{ background: #fff7e6; color: #d48806; }}
        .badge-danger {{ background: #fff1f0; color: #cf1322; }}
        .badge-info {{ background: #e6f4ff; color: #0958d9; }}

        .hub-tag {{
            display: inline-block;
            padding: 2px 8px;
            background: #e6f4ff;
            color: #0958d9;
            border-radius: 4px;
            font-size: 12px;
            margin-right: 5px;
        }}
        .site-tag {{
            display: inline-block;
            padding: 2px 8px;
            background: #f6ffed;
            color: #389e0d;
            border-radius: 4px;
            font-size: 12px;
            margin-right: 5px;
        }}

        .grand-total td {{
            background: #1a3a5c !important;
            color: white !important;
            font-weight: bold;
            font-size: 15px;
            padding: 16px 14px;
        }}

        .note {{
            margin-bottom: 20px;
            padding: 12px 14px;
            background: #fffbe6;
            border-left: 4px solid #faad14;
            border-radius: 4px;
            font-size: 12px;
            color: #888;
            line-height: 1.6;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>美西OTWS门户推广数据分析报表-0608-0614</h1>

        <div class="summary">
            <div class="rate-row">
                <div class="section-label">关键指标</div>
                <div class="card">
                    <div class="card-label">总体需求满足率</div>
                    <div class="card-value">{total_fulfill_rate:.1f}%</div>
                </div>
                <div class="card">
                    <div class="card-label">总体线上派遣率</div>
                    <div class="card-value">{total_online_rate:.1f}%</div>
                </div>
            </div>
            <div class="number-row">
                <div class="section-label">发单及派遣情况</div>
                <div class="card">
                    <div class="card-label">总需求数(发单数)</div>
                    <div class="card-value">{total_demand:,}</div>
                </div>
                <div class="card">
                    <div class="card-label">实际线上派遣</div>
                    <div class="card-value">{total_online:,}</div>
                </div>
                <div class="card">
                    <div class="card-label">线下派遣数</div>
                    <div class="card-value">{total_offline:,}</div>
                </div>
            </div>
            <div class="area-row">
                <div class="section-label">区域情况</div>
                <div class="card">
                    <div class="card-label">供应商数量</div>
                    <div class="card-value">{len(all_suppliers)}</div>
                </div>
                <div class="card">
                    <div class="card-label">HUB数量</div>
                    <div class="card-value">{len(hubs)}</div>
                </div>
                <div class="card">
                    <div class="card-label">站点数量</div>
                    <div class="card-value">{len(sites)}</div>
                </div>
            </div>
        </div>

        <div class="note">
            <strong>说明：</strong>
            <ul style="margin: 10px 0 0 20px; line-height: 1.8;">
                <li>需求满足率 = 实际线上派遣数 / 发单需求数</li>
                <li>线上派遣率 = 1 − 线下派遣数 / (线下派遣数 + 实际线上派遣数)</li>
            </ul>
        </div>

        <table>
            <thead>
                <tr>
                    <th>需求仓</th>
                    <th>供应商</th>
                    <th>类型</th>
                    <th>发单需求数</th>
                    <th>需求满足率</th>
                    <th>实际线上派遣</th>
                    <th>线下派遣数</th>
                    <th>线上派遣率</th>
                </tr>
            </thead>
            <tbody>
'''

    for wh, data in sorted_warehouses:
        suppliers_list = data['suppliers']
        count = len(suppliers_list)

        for sp, demand, online, offline in suppliers_list:
            # 需求满足率 = 实际线上派遣数 / 发单需求数
            if demand == 0:
                fr_display = '<span class="badge badge-info">/</span>'
            else:
                fulfill_rate = online / demand * 100
                if fulfill_rate >= 90:
                    fr_badge = 'badge-success'
                elif fulfill_rate >= 70:
                    fr_badge = 'badge-warning'
                else:
                    fr_badge = 'badge-danger'
                fr_display = f'<span class="badge {fr_badge}">{fulfill_rate:.1f}%</span>'

            # 线上派遣率 = 1 - 线下派遣数 / (线下派遣数 + 实际线上派遣数)
            total_dispatched = online + offline
            if total_dispatched == 0:
                or_display = '<span class="badge badge-info">/</span>'
            else:
                online_rate = online / total_dispatched * 100
                if online_rate >= 90:
                    or_display = f'<span class="badge badge-success">{online_rate:.1f}%</span>'
                elif online_rate >= 70:
                    or_display = f'<span class="badge badge-warning">{online_rate:.1f}%</span>'
                else:
                    or_display = f'<span class="badge badge-danger">{online_rate:.1f}%</span>'

            html += f'''                <tr>
                    <td><strong>{wh}</strong></td>
                    <td><strong>{sp}</strong></td>
                    <td><span class="{'hub-tag' if '.H' in wh else 'site-tag'}">{'HUB' if '.H' in wh else '站点'}</span></td>
                    <td class="num">{demand:,}</td>
                    <td class="rate">{fr_display}</td>
                    <td class="num">{online:,}</td>
                    <td class="num">{offline:,}</td>
                    <td class="rate">{or_display}</td>
                </tr>
'''

        # Subtotal row (only if more than 1 supplier)
        if count > 1:
            wh_total_d = data['total_demand']
            wh_total_on = data['total_online']
            wh_total_off = data.get('total_offline', 0)
            wh_fulfill = wh_total_on / wh_total_d * 100 if wh_total_d > 0 else 0
            wh_on_rate = wh_total_on / (wh_total_on + wh_total_off) * 100 if (wh_total_on + wh_total_off) > 0 else 0
            html += f'''                <tr style="background:#fafafa;font-weight:bold;border-top:2px solid #d9d9d9;">
                    <td style="color:#d48806;">{wh} 小计</td>
                    <td></td>
                    <td></td>
                    <td class="num">{wh_total_d:,}</td>
                    <td class="rate">{wh_fulfill:.1f}%</td>
                    <td class="num">{wh_total_on:,}</td>
                    <td class="num">{wh_total_off:,}</td>
                    <td class="rate">{wh_on_rate:.1f}%</td>
                </tr>
'''

    # Grand total row
    html += f'''                <tr class="grand-total">
                    <td>总计</td>
                    <td colspan="2"></td>
                    <td class="num">{total_demand:,}</td>
                    <td class="rate">{total_fulfill_rate:.1f}%</td>
                    <td class="num">{total_online:,}</td>
                    <td class="num">{total_offline:,}</td>
                    <td class="rate">{total_online_rate:.1f}%</td>
                </tr>
'''

    html += '''            </tbody>
        </table>
    </div>
</body>
</html>
'''

    output_path = r'C:\Users\Administrator\WorkBuddy\2026-06-16-19-18-38\美西OTWS门户推广数据分析报表_供应商维度_0608-0614.html'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"HTML 已生成: {output_path}")
    return output_path


def generate_excel(by_warehouse, total_offline):
    """生成需求仓维度Excel报表"""
    import xlsxwriter

    output_path = r'C:\Users\Administrator\WorkBuddy\2026-06-16-19-18-38\美西OTWS门户推广数据分析报表_供应商维度_0608-0614.xlsx'
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet('需求仓维度分析')

    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#1a3a5c', 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    cell_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    num_format = workbook.add_format({
        'border': 1, 'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0'
    })
    pct_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.0"%"'
    })
    subtotal_format = workbook.add_format({
        'bold': True, 'bg_color': '#fafafa', 'border': 1,
        'align': 'center', 'valign': 'vcenter', 'top': 2
    })
    subtotal_num_format = workbook.add_format({
        'bold': True, 'bg_color': '#fafafa', 'border': 1,
        'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0', 'top': 2
    })
    subtotal_pct_format = workbook.add_format({
        'bold': True, 'bg_color': '#fafafa', 'border': 1,
        'align': 'center', 'valign': 'vcenter', 'num_format': '0.0"%"', 'top': 2
    })

    grand_total_format = workbook.add_format({
        'bold': True, 'bg_color': '#1a3a5c', 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14
    })
    grand_total_num_format = workbook.add_format({
        'bold': True, 'bg_color': '#1a3a5c', 'font_color': 'white',
        'border': 1, 'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0', 'font_size': 14
    })
    grand_total_pct_format = workbook.add_format({
        'bold': True, 'bg_color': '#1a3a5c', 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.0"%"', 'font_size': 14
    })

    headers = ['需求仓', '供应商', '类型', '发单需求数', '需求满足率', '实际线上派遣', '线下派遣数', '线上派遣率']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    # Sort warehouses: HUBs first by order, then sites
    hub_order = ["CNO.H", "LAV.H", "SFO.H", "SEA.H"]
    site_order = ["DEN", "PHX", "SAN01"]

    def wh_sort_key(wh):
        if wh in hub_order:
            return (0, hub_order.index(wh))
        elif wh in site_order:
            return (1, site_order.index(wh))
        else:
            return (2, wh)

    sorted_warehouses = sorted(by_warehouse.items(), key=lambda x: wh_sort_key(x[0]))

    row = 1
    total_d_all = 0
    total_on_all = 0
    total_off_all = 0

    for wh, data in sorted_warehouses:
        suppliers_list = data['suppliers']
        count = len(suppliers_list)
        wh_total_d = data['total_demand']
        wh_total_on = data['total_online']
        wh_total_off = data['total_offline']

        total_d_all += wh_total_d
        total_on_all += wh_total_on
        total_off_all += wh_total_off

        for sp, demand, online, offline in suppliers_list:
            worksheet.write(row, 0, wh, cell_format)
            worksheet.write(row, 1, sp, cell_format)
            worksheet.write(row, 2, 'HUB' if '.H' in wh else '站点', cell_format)
            worksheet.write(row, 3, demand, num_format)
            if demand == 0:
                worksheet.write(row, 4, '/', cell_format)
            else:
                fulfill_rate = online / demand
                worksheet.write(row, 4, fulfill_rate, pct_format)
            # 线上派遣率 = 1 - 线下派遣数/(线下派遣数+实际线上派遣数)
            total_dispatched = online + offline
            if total_dispatched == 0:
                worksheet.write(row, 7, '/', cell_format)
            else:
                online_rate = online / total_dispatched
                worksheet.write(row, 7, online_rate, pct_format)
            worksheet.write(row, 5, online, num_format)
            worksheet.write(row, 6, offline, num_format)
            row += 1

        # Subtotal row (only if > 1 supplier)
        if count > 1:
            wh_fulfill = wh_total_on / wh_total_d if wh_total_d > 0 else 0
            wh_on_rate = wh_total_on / (wh_total_on + wh_total_off) if (wh_total_on + wh_total_off) > 0 else 0
            worksheet.write(row, 0, f'{wh} 小计', subtotal_format)
            worksheet.write(row, 1, '', subtotal_format)
            worksheet.write(row, 2, '', subtotal_format)
            worksheet.write(row, 3, wh_total_d, subtotal_num_format)
            worksheet.write(row, 4, wh_fulfill, subtotal_pct_format)
            worksheet.write(row, 5, wh_total_on, subtotal_num_format)
            worksheet.write(row, 6, wh_total_off, subtotal_num_format)
            worksheet.write(row, 7, wh_on_rate, subtotal_pct_format)
            row += 1

    # Grand total
    total_fulfill_all = total_on_all / total_d_all if total_d_all > 0 else 0
    total_online_all_rate = total_on_all / (total_on_all + total_off_all) if (total_on_all + total_off_all) > 0 else 0

    worksheet.merge_range(row, 0, row, 1, '总计', grand_total_format)
    worksheet.write(row, 2, '', grand_total_format)
    worksheet.write(row, 3, total_d_all, grand_total_num_format)
    worksheet.write(row, 4, total_fulfill_all, grand_total_pct_format)
    worksheet.write(row, 5, total_on_all, grand_total_num_format)
    worksheet.write(row, 6, total_off_all, grand_total_num_format)
    worksheet.write(row, 7, total_online_all_rate, grand_total_pct_format)

    worksheet.set_column(0, 0, 12)
    worksheet.set_column(1, 1, 22)
    worksheet.set_column(2, 2, 8)
    worksheet.set_column(3, 3, 16)
    worksheet.set_column(4, 4, 14)
    worksheet.set_column(5, 5, 18)
    worksheet.set_column(6, 6, 18)
    worksheet.set_column(7, 7, 14)

    workbook.close()
    print(f"Excel 已生成: {output_path}")
    return output_path


def main():
    print("=" * 60)
    print("(0608-0614)美西OTWS门户推广数据分析报表 — 需求仓维度")
    print("=" * 60)

    print("\n构建需求仓维度数据...")
    by_warehouse, all_suppliers, total_offline = build_data()

    print(f"\n共 {len(by_warehouse)} 个需求仓, {len(all_suppliers)} 家供应商:")

    hub_order = ["CNO.H", "LAV.H", "SFO.H", "SEA.H"]
    site_order = ["DEN", "PHX", "SAN01"]
    def wh_sort_key(wh):
        if wh in hub_order:
            return (0, hub_order.index(wh))
        elif wh in site_order:
            return (1, site_order.index(wh))
        else:
            return (2, wh)
    sorted_warehouses = sorted(by_warehouse.items(), key=lambda x: wh_sort_key(x[0]))

    for wh, data in sorted_warehouses:
        sps = ', '.join(f"{sp}({d})" for sp, d, _, _ in data['suppliers'])
        print(f"  {wh}: 总需求={data['total_demand']}, 线上={data['total_online']}, 线下={data['total_offline']}")
        print(f"    供应商: {sps}")

    print("\n生成HTML报表...")
    html_path = generate_html(by_warehouse, all_suppliers, total_offline)

    print("生成Excel报表...")
    excel_path = generate_excel(by_warehouse, total_offline)

    print("\n报表生成完成！")


if __name__ == '__main__':
    main()

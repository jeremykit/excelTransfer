import os
from io import BytesIO
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
import warnings

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

warnings.simplefilter(action='ignore', category=FutureWarning)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files: return jsonify({'error': '无文件'}), 400
    file = request.files['file']
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'temp.xlsx')
    file.save(filepath)
    try:
        xl = pd.ExcelFile(filepath)
        return jsonify({'sheets': xl.sheet_names})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/preview_sheet', methods=['POST'])
def preview_sheet():
    data = request.json
    sheet_name = data.get('sheet_name')
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'temp.xlsx')
    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=20)
        df = df.fillna('')
        return jsonify({ 'rows': df.values.tolist(), 'col_count': df.shape[1] })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/transform', methods=['POST'])
def transform():
    data = request.json
    sheet_name = data.get('sheet_name')
    # header_row_index 是用户点击的那一行，数据从它下一行开始
    header_idx = int(data.get('header_row_index', 0))
    mapping = data.get('mapping')  # 现在的 mapping value 是列索引 (0, 1, 14...)
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'temp.xlsx')
    
    try:
        # 1. 即使有表头，我们也用 header=None 读，然后手动切片，这样最稳，完全通过 index 操作
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        
        # 数据从 header_idx + 1 开始
        df_data = df.iloc[header_idx + 1:].copy()

        # 准备原始数据展示
        raw_display = []

        # 准备清洗后数据
        clean_rows = []

        # 获取用户映射的列索引 (转为 int)
        idx_dev_name = int(mapping.get('device_group')) if mapping.get('device_group') is not None else None
        idx_point = int(mapping.get('point_name')) if mapping.get('point_name') is not None else None

        # === 核心清洗逻辑 ===

        # A. 空值填充 (Fill Down)
        if idx_dev_name is not None:
            df_data[idx_dev_name] = df_data[idx_dev_name].infer_objects(copy=False).ffill()

        for _, row in df_data.iterrows():
            if idx_point is not None and pd.isna(row[idx_point]):
                continue

            raw_item = {}
            if idx_dev_name is not None:
                raw_item['dev'] = str(row[idx_dev_name]) if pd.notna(row[idx_dev_name]) else ''
            if idx_point is not None:
                raw_item['pt'] = str(row[idx_point]) if pd.notna(row[idx_point]) else ''
            sig_idx = int(mapping.get('signal_type')) if mapping.get('signal_type') is not None else None
            if sig_idx is not None:
                raw_item['sig'] = str(row[sig_idx]) if pd.notna(row[sig_idx]) else ''
            raw_display.append(raw_item)

            item = {
                'device_group': raw_item.get('dev', ''),
                'project_no': '',
                'name_en': '',
                'name_zh': '',
                'alias': '',
                'data_type': '',
                'signal_type': '',
                'business_type': '',
                'alarm_type': '',
                'status_enum': '',
                'collect_type': '',
                'related_point': '',
                'range': '',
                'unit': '',
                'll': '',
                'l': '',
                'h': '',
                'hh': ''
            }

            for std_key, col_idx_str in mapping.items():
                if col_idx_str is None or col_idx_str == "":
                    continue
                col_idx = int(col_idx_str)

                if std_key == 'point_name':
                    val = str(row[col_idx]) if pd.notna(row[col_idx]) else ""
                    if '\n' in val:
                        parts = val.split('\n')
                        item['name_en'] = parts[0].strip()
                        item['name_zh'] = parts[1].strip() if len(parts) > 1 else ''
                    else:
                        item['name_en'] = val
                        item['name_zh'] = val
                elif std_key == 'device_group':
                    item['device_group'] = str(row[col_idx]) if pd.notna(row[col_idx]) else ""
                elif std_key == 'item_no':
                    item['project_no'] = str(row[col_idx]) if pd.notna(row[col_idx]) else ""
                elif std_key == 'signal_type':
                    raw_sig = str(row[col_idx]) if pd.notna(row[col_idx]) else ""
                    item['signal_type'] = raw_sig
                    normalized = raw_sig.lower()
                    if 'analog' in normalized or '模拟' in raw_sig:
                        item['data_type'] = 'DOUBLE'
                        item['business_type'] = 'DIGITAL'
                    if 'on' in normalized or 'off' in normalized or 'switch' in normalized or '开关' in raw_sig:
                        item['data_type'] = 'DOUBLE'
                        item['business_type'] = 'STATE'
                elif std_key == 'alarm_type':
                    val = str(row[col_idx]) if pd.notna(row[col_idx]) else ""
                    item['alarm_type'] = val
                    if val:
                        item['business_type'] = 'ALARM'
                else:
                    item[std_key] = str(row[col_idx]) if pd.notna(row[col_idx]) else ""

            clean_rows.append(item)

        return jsonify({
            'raw_data': raw_display,
            'clean_data': clean_rows
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)


@app.route('/api/export', methods=['POST'])
def export_file():
    payload = request.json
    rows = payload.get('rows', [])
    ship_info = payload.get('ship_info', {})

    if not rows:
        return jsonify({'error': '没有可导出的数据'}), 400

    df_points = pd.DataFrame(rows)
    if df_points.empty:
        return jsonify({'error': '数据为空'}), 400

    df_points['group_name'] = df_points.get('group_name', '').replace('', '默认分组')
    df_points['device_name'] = df_points.get('device_group', '').replace('', '未命名设备')

    group_ids = {name: idx + 1 for idx, name in enumerate(df_points['group_name'].unique())}
    device_ids = {name: idx + 1 for idx, name in enumerate(df_points['device_name'].unique())}

    df_points['group_id'] = df_points['group_name'].map(group_ids)
    df_points['device_id'] = df_points['device_name'].map(device_ids)

    ship_df = pd.DataFrame([{
        '船只名称(Name)': ship_info.get('name', ''),
        '船只编号(HULL No.)': ship_info.get('hull', ''),
        '船东(Owner)': ship_info.get('owner', ''),
        '项目编号(Project No.)': ship_info.get('project', ''),
        '船级(Class)': ship_info.get('class', ''),
        'IMO编号': ship_info.get('imo', ''),
        'MMSI编号': ship_info.get('mmsi', '')
    }])

    group_df = pd.DataFrame([{
        'ID': gid,
        'name_en': name,
        'name_zh': name,
        'alias': ''
    } for name, gid in group_ids.items()])

    device_df = pd.DataFrame([{
        'id': did,
        'name_en': name,
        'name_zh': name,
        'alias': '',
        '类别': '',
        'group_id': df_points[df_points['device_name'] == name]['group_id'].iloc[0],
        'product_name': '',
        'IP地址': df_points[df_points['device_name'] == name]['ip_address'].iloc[0] if 'ip_address' in df_points.columns else ''
    } for name, did in device_ids.items()])

    point_columns = [
        '设备名', '项目编号', '标准名称', '检测点名称(EN)', '检测点名称(ZH/ITEM NAME)', '别名',
        '数据类型', '信号类型', '业务类型', '报警类型', '状态枚举', '采集类型', '关联测点',
        '范围', '单位', '低低限', '低限', '高限', '高高限'
    ]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        ship_df.to_excel(writer, sheet_name='说明', index=False, startrow=0)
        group_df.to_excel(writer, sheet_name='说明', index=False, startrow=3)
        device_df.to_excel(writer, sheet_name='说明', index=False, startrow=8)

        for group_name, group_data in df_points.groupby('group_name'):
            sheet_rows = []
            for device_name, device_data in group_data.groupby('device_name'):
                for _, r in device_data.iterrows():
                    sheet_rows.append({
                        '设备名': device_name,
                        '项目编号': r.get('project_no', ''),
                        '标准名称': r.get('standard_name', ''),
                        '检测点名称(EN)': r.get('name_en', ''),
                        '检测点名称(ZH/ITEM NAME)': r.get('name_zh', ''),
                        '别名': r.get('alias', ''),
                        '数据类型': r.get('data_type', ''),
                        '信号类型': r.get('signal_type', ''),
                        '业务类型': r.get('business_type', ''),
                        '报警类型': r.get('alarm_type', ''),
                        '状态枚举': r.get('status_enum', ''),
                        '采集类型': r.get('collect_type', ''),
                        '关联测点': r.get('related_point', ''),
                        '范围': r.get('range', ''),
                        '单位': r.get('unit', ''),
                        '低低限': r.get('ll', ''),
                        '低限': r.get('l', ''),
                        '高限': r.get('h', ''),
                        '高高限': r.get('hh', '')
                    })
                sheet_rows.append({})

            pd.DataFrame(sheet_rows, columns=point_columns).to_excel(
                writer, sheet_name=group_name or '分组', index=False
            )

    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='standard_points.xlsx'
    )

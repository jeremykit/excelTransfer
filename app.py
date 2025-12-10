import os
import pandas as pd
from flask import Flask, render_template, request, jsonify
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
    mapping = data.get('mapping') # 现在的 mapping value 是列索引 (0, 1, 14...)
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'temp.xlsx')
    
    try:
        # 1. 即使有表头，我们也用 header=None 读，然后手动切片，这样最稳，完全通过 index 操作
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        
        # 数据从 header_idx + 1 开始
        df_data = df.iloc[header_idx + 1 : header_idx + 51].copy() # 取50行预览
        
        # 准备原始数据展示
        raw_display = []
        
        # 准备清洗后数据
        clean_rows = []
        
        # 获取用户映射的列索引 (转为 int)
        idx_dev_grp = int(mapping.get('device_group')) if mapping.get('device_group') is not None else None
        idx_point = int(mapping.get('point_name')) if mapping.get('point_name') is not None else None
        
        # === 核心清洗逻辑 ===
        
        # A. 空值填充 (Fill Down)
        # 注意：这里我们是对某一列进行填充
        if idx_dev_grp is not None:
            # infer_objects 消除警告
            df_data[idx_dev_grp] = df_data[idx_dev_grp].infer_objects(copy=False).ffill()

        for _, row in df_data.iterrows():
            # 跳过完全无效行 (比如测点名是空的)
            if idx_point is not None and pd.isna(row[idx_point]):
                continue

            # 1. 构建原始数据行 (仅用于前端展示对比)
            raw_item = {}
            if idx_dev_grp is not None: raw_item['dev'] = str(row[idx_dev_grp])
            if idx_point is not None: raw_item['pt'] = str(row[idx_point])
            # 其他随便取一个做展示
            sig_idx = int(mapping.get('signal_type')) if mapping.get('signal_type') is not None else None
            if sig_idx is not None: raw_item['sig'] = str(row[sig_idx])
            raw_display.append(raw_item)

            # 2. 构建清洗后数据
            item = {}
            # 通用字段直接取值
            for std_key, col_idx_str in mapping.items():
                if col_idx_str is None or col_idx_str == "": continue
                col_idx = int(col_idx_str)
                
                # 特殊处理：测点名拆分
                if std_key == 'point_name':
                    val = str(row[col_idx]) if pd.notna(row[col_idx]) else ""
                    if '\n' in val:
                        parts = val.split('\n')
                        item['name_en'] = parts[0].strip()
                        item['name_zh'] = parts[1].strip() if len(parts) > 1 else ''
                    else:
                        item['name_en'] = val
                        item['name_zh'] = val
                # 特殊处理：设备名 (已经 Fill Down 过了)
                elif std_key == 'device_group':
                    item['device_group'] = str(row[col_idx]) if pd.notna(row[col_idx]) else ""
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

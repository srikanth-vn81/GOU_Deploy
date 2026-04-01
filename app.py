from flask import Flask, request, jsonify, send_file
import pandas as pd
import os

app = Flask(__name__, static_folder='.', static_url_path='')

os.makedirs('uploads', exist_ok=True)

@app.route('/')
def index():
    return app.send_static_file('index.html')


# ── ROUTE 1: GOU REPORT PROCESSOR ────────────────────
@app.route('/process', methods=['POST'])
def process():
    file = request.files['file']
    input_path  = 'uploads/' + file.filename
    output_path = r"C:\GOU\GOU_final.xlsx"
    file.save(input_path)

    engine = 'pyxlsb' if file.filename.endswith('.xlsb') else 'openpyxl'
    df = pd.read_excel(input_path, sheet_name='Client List', engine=engine, skiprows=1)

    drop_cols = ['Actionable ( Yes/N0)', 'Red Flags', 'BILLING_METHOD', 'Code', 'Status', 'Final']
    df.drop(columns=[c for c in drop_cols if c in df.columns], inplace=True)

    df.rename(columns={
        'Jan-26 to Apr-26':    'Req HC',
        'Jan-26 to Apr-26.1':  'Fcst HC',
        'Staffing - Reason':   'Staffing - Reason 1',
        'Bucket':              'Bucket - 1',
        'Commercial impact $': 'Commercial impact ',
    }, inplace=True)

    staffing_val = df['Staffing Category'] if 'Staffing Category' in df.columns else None

    inserts = [
        (11, 'Staffing Stage',                staffing_val),
        (16, 'Bucket-1 $',                    None),
        (17, 'Staffing - Reason 2',           None),
        (18, 'Bucket - 2',                    None),
        (20, 'Bucket-2 $',                    None),
        (21, 'Staffing - Reason 3',           None),
        (22, 'Bucket - 3',                    None),
        (24, 'Bucket-3 $',                    None),
        (28, 'Business Impact',               None),
        (29, 'Commercial impact $ (Revenue)', None),
        (31, 'Revenue By HC',                 None),
    ]
    for pos, col_name, val in inserts:
        try:
            df.insert(loc=pos, column=col_name, value=val)
        except Exception as e:
            print(f"Could not insert '{col_name}': {e}")

    df.to_excel(output_path, index=False, sheet_name='Client List')

    df2 = pd.read_excel(output_path)
    df2.rename(columns={df2.columns[5]: 'Req HC', df2.columns[6]: 'Fcst HC'}, inplace=True)
    df2.to_excel(output_path, index=False)

    return jsonify({'success': True, 'msg': f"✅ Saved to {output_path}"})


# ── ROUTE 2: HC PIE ───────────────────────────────────
@app.route('/run-hc-pie', methods=['POST'])
def run_hc_pie():
    try:
        input_path  = r"C:\GOU\GOU_final.xlsx"
        output_path = r"C:\GOU\hc_pie.xlsx"

        df = pd.read_excel(input_path, sheet_name="Sheet1")
        df_crossskill = df[df["CROSS SKILLING - CHECK"] == 9].copy()
        df_crossskill["Req HC"] = pd.to_numeric(df_crossskill["Req HC"], errors="coerce").fillna(0)

        df_hc = (
            df_crossskill
            .groupby("Staffing Category", dropna=False)["Req HC"]
            .sum()
            .reset_index()
        )

        df_hc.to_excel(output_path, index=False)
        return jsonify({'success': True, 'msg': f"✅ {len(df_hc)} rows saved → {output_path}"})

    except Exception as e:
        return jsonify({'success': False, 'msg': f"❌ Error: {str(e)}"})


# ── ROUTE 3: LOB PIE ──────────────────────────────────
@app.route('/run-lob-pie', methods=['POST'])
def run_lob_pie():
    try:
        input_path  = r"C:\GOU\GOU_final.xlsx"
        output_path = r"C:\GOU\lob_pie.xlsx"

        df = pd.read_excel(input_path, sheet_name="Sheet1")
        df_crossskill = df[df["CROSS SKILLING - CHECK"] == 9].copy()

        df_lob = (
            df_crossskill
            .groupby("Staffing Category")
            .size()
            .reset_index(name="Column1")
        )

        df_lob.to_excel(output_path, index=False)
        return jsonify({'success': True, 'msg': f"✅ {len(df_lob)} rows saved → {output_path}"})

    except Exception as e:
        return jsonify({'success': False, 'msg': f"❌ Error: {str(e)}"})


# ── ROUTE 4: DOWNLOAD BY PATH ─────────────────────────
@app.route('/download-path')
def download_path():
    path = request.args.get('path')
    if not path or not os.path.exists(path):
        return jsonify({'error': 'File not found'}), 404
    return send_file(path, as_attachment=True)


# ── ROUTE 5: DOWNLOAD FROM C:\GOU\ ───────────────────
@app.route('/download/<filename>')
def download(filename):
    path = r"C:\GOU\\" + filename
    if not os.path.exists(path):
        return jsonify({'error': 'File not found'}), 404
    return send_file(path, as_attachment=True)


if __name__ == '__main__':
    print("Open → http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
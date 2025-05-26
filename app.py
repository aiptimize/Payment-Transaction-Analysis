from flask import Flask, request, render_template, redirect, url_for, session
from werkzeug.utils import secure_filename
import os
import pdfplumber
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
import matplotlib
matplotlib.use('Agg')  # 设置为非交互式后端
import matplotlib.pyplot as plt

app = Flask(__name__)
app.secret_key = 'your_secret_key'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['CHART_FOLDER'] = 'charts'

DEFAULT_SENSITIVE_WORDS = ['法院', '银行']
DEFAULT_INDUSTRY_KEYWORDS = ['水泥', '钢材']

# 设置matplotlib字体
plt.rcParams['font.family'] = 'SimHei'  # 使用黑体字体，可根据系统中实际字体修改
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题


def extract_location(text, keywords):
    for keyword in keywords:
        if keyword in str(text):
            return keyword
    return "其他"


def add_text_to_doc(text, font_size=10, bold=False):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold


def process_wechat_pdf(input_pdf_path, output_doc_path, sensitive_words, industry_keywords):
    try:
        if not os.path.exists(input_pdf_path):
            raise FileNotFoundError(f"输入的PDF文件不存在: {input_pdf_path}")

        all_data = []
        header = None

        with pdfplumber.open(input_pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                table = page.extract_table()
                if table:
                    if i == 0:
                        header = table[0]
                        all_data.extend(table[1:])
                    else:
                        all_data.extend(table[1:] if table[0] == header else table)

        if all_data:
            all_data = all_data[2:]
            new_columns = [
                "交易单号",
                "交易时间",
                "交易类型",
                "收/支/其他",
                "交易方式",
                "金额(元)",
                "交易对方",
                "商户单号"
            ]
            df = pd.DataFrame(all_data, columns=new_columns)
        else:
            raise ValueError("未找到表格数据")

        df = df.apply(lambda x: x.map(lambda y: y.replace('\n', '') if isinstance(y, str) else y))
        df['金额(元)'] = pd.to_numeric(df['金额(元)'], errors='coerce')

        global doc
        doc = Document()
        add_text_to_doc("微信支付交易明细分析报告", font_size=16, bold=True)
        add_text_to_doc("-" * 80)
        doc.add_paragraph()

        doc.add_heading('目录', level=1)
        doc.add_paragraph('1. 敏感词消费记录\t\t\t\t页 1')
        doc.add_paragraph('2. 行业消费记录\t\t\t\t页 2')
        doc.add_paragraph('3. 区域消费金额统计\t\t\t\t页 3')
        doc.add_paragraph('4. 交易金额前5的交易对方\t\t\t页 4')
        doc.add_paragraph('5. 交易数量前5的交易对方\t\t\t页 5')
        doc.add_page_break()

        filtered_df = df[df["交易对方"].str.contains('|'.join(sensitive_words), na=False)].copy()
        columns_to_drop = ['交易单号', '商户单号']
        for col in columns_to_drop:
            if col in filtered_df.columns:
                filtered_df = filtered_df.drop(columns=[col])

        if not filtered_df.empty:
            sorted_df = filtered_df.dropna(subset=["金额(元)"]).sort_values(by="金额(元)", ascending=False)
            sorted_df.reset_index(drop=True, inplace=True)
            doc.add_heading('1. 敏感词消费记录', level=1)
            add_text_to_doc("敏感词消费记录（按金额降序排列）：", font_size=12, bold=True)
            table = doc.add_table(rows=1, cols=len(sorted_df.columns))
            hdr_cells = table.rows[0].cells
            for i, col in enumerate(sorted_df.columns):
                hdr_cells[i].text = col
            for _, row in sorted_df.head(5).iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row):
                    row_cells[i].text = str(value)
        else:
            doc.add_heading('1. 敏感词消费记录', level=1)
            add_text_to_doc("没有敏感词消费记录")

        doc.add_page_break()

        filtered_df = df[df["交易对方"].str.contains('|'.join(industry_keywords), na=False, case=False)].copy()
        for col in columns_to_drop:
            if col in filtered_df.columns:
                filtered_df = filtered_df.drop(columns=[col])

        if not filtered_df.empty:
            sorted_df = filtered_df.dropna(subset=["金额(元)"]).sort_values(by="金额(元)", ascending=False)
            sorted_df.reset_index(drop=True, inplace=True)
            doc.add_heading('2. 行业消费记录', level=1)
            add_text_to_doc("行业消费记录（按金额降序排列）：", font_size=12, bold=True)
            table = doc.add_table(rows=1, cols=len(sorted_df.columns))
            hdr_cells = table.rows[0].cells
            for i, col in enumerate(sorted_df.columns):
                hdr_cells[i].text = col
            for _, row in sorted_df.head(5).iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row):
                    row_cells[i].text = str(value)
        else:
            doc.add_heading('2. 行业消费记录', level=1)
            add_text_to_doc("没有行业消费记录")

        doc.add_page_break()

        location_keywords = ["潢川县", "深圳市", "郑州市", "温县"]
        column_variants = ['交易对方', '对方', '收款方', '商户名称', 'counterpart', 'recipient']
        target_column = next((col for col in column_variants if col in df.columns), None)

        if target_column is None:
            available_cols = ", ".join(df.columns.tolist())
            raise KeyError(f"找不到交易对方列，可用的列有：{available_cols}")

        filtered_df = df[df[target_column].str.contains('|'.join(location_keywords), na=False, case=False, regex=True)].copy()
        amount_variants = ['金额(元)', '金额', 'money', 'amount', '交易金额']
        amount_column = next((col for col in amount_variants if col in df.columns), None)

        if amount_column is None:
            available_cols = ", ".join(df.columns.tolist())
            raise KeyError(f"找不到金额列，可用的列有：{available_cols}")

        sorted_df = filtered_df.dropna(subset=[amount_column]).sort_values(by=amount_column, ascending=False)
        pd.set_option('display.max_rows', None)
        pd.set_option('display.unicode.ambiguous_as_wide', True)
        pd.set_option('display.unicode.east_asian_width', True)

        sorted_df['区域'] = sorted_df[target_column].apply(lambda x: extract_location(x, location_keywords))
        region_stats = sorted_df.groupby('区域')[amount_column].sum().sort_values(ascending=False).head(5)

        doc.add_heading('3. 区域消费金额统计', level=1)
        add_text_to_doc("区域消费金额统计：", font_size=12, bold=True)
        table = doc.add_table(rows=1, cols=2)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '区域'
        hdr_cells[1].text = '金额(元)'
        for region, amount in region_stats.items():
            row_cells = table.add_row().cells
            row_cells[0].text = region
            row_cells[1].text = f"{amount:.2f}"

        plt.figure(figsize=(10, 6))
        plt.bar(region_stats.index, region_stats.values)
        plt.xlabel('区域')
        plt.ylabel('金额(元)')
        plt.title('区域消费金额统计')
        chart_path = os.path.join(app.config['CHART_FOLDER'], 'region_chart.png')
        plt.savefig(chart_path)
        plt.close()
        doc.add_picture(chart_path, width=Inches(6))

        doc.add_page_break()

        top_5_by_amount = df.groupby('交易对方')['金额(元)'].sum().abs().sort_values(ascending=False).head(5)
        top_5_by_count = df.groupby('交易对方')['金额(元)'].count().sort_values(ascending=False).head(5)

        doc.add_heading('4. 交易金额前5的交易对方', level=1)
        add_text_to_doc('交易金额前5的交易对方：', font_size=12, bold=True)
        table = doc.add_table(rows=1, cols=2)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '交易对方'
        hdr_cells[1].text = '金额(元)'
        for entity, amount in top_5_by_amount.items():
            row_cells = table.add_row().cells
            row_cells[0].text = entity
            row_cells[1].text = f"{amount:.2f}"

        plt.figure(figsize=(10, 6))
        plt.bar(top_5_by_amount.index, top_5_by_amount.values)
        plt.xlabel('交易对方')
        plt.ylabel('金额(元)')
        plt.title('交易金额前5的交易对方')
        chart_path = os.path.join(app.config['CHART_FOLDER'], 'top_5_amount_chart.png')
        plt.savefig(chart_path)
        plt.close()
        doc.add_picture(chart_path, width=Inches(6))

        doc.add_page_break()

        doc.add_heading('5. 交易数量前5的交易对方', level=1)
        add_text_to_doc('交易数量前5的交易对方：', font_size=12, bold=True)
        table = doc.add_table(rows=1, cols=2)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '交易对方'
        hdr_cells[1].text = '交易数量'
        for entity, count in top_5_by_count.items():
            row_cells = table.add_row().cells
            row_cells[0].text = entity
            row_cells[1].text = str(count)

        plt.figure(figsize=(10, 6))
        plt.bar(top_5_by_count.index, top_5_by_count.values)
        plt.xlabel('交易对方')
        plt.ylabel('交易数量')
        plt.title('交易数量前5的交易对方')
        chart_path = os.path.join(app.config['CHART_FOLDER'], 'top_5_count_chart.png')
        plt.savefig(chart_path)
        plt.close()
        doc.add_picture(chart_path, width=Inches(6))

        doc.save(output_doc_path)

        if not os.path.exists(output_doc_path):
            raise Exception(f"DOCX保存失败，文件不存在: {output_doc_path}")

    except Exception as e:
        if os.path.exists(output_doc_path):
            try:
                os.remove(output_doc_path)
            except:
                pass
        raise


def process_alipay_pdf(input_pdf_path, output_doc_path, sensitive_words, industry_keywords):
    try:
        if not os.path.exists(input_pdf_path):
            raise FileNotFoundError(f"输入的PDF文件不存在: {input_pdf_path}")

        all_data = []
        header = None

        with pdfplumber.open(input_pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                table = page.extract_table()
                if table:
                    if i == 0:
                        header = table[0]
                        all_data.extend(table[1:])
                    else:
                        all_data.extend(table[1:] if table[0] == header else table)

        if all_data:
            all_data = all_data[2:]
            new_columns = [
                "收/支",
                "交易对方",
                "商品说明",
                "收/付款方式",
                "金额",
                "交易订单号",
                "商家订单号",
                "交易时间"
            ]
            df = pd.DataFrame(all_data, columns=new_columns)
        else:
            raise ValueError("未找到表格数据")

        df = df.apply(lambda x: x.map(lambda y: y.replace('\n', '') if isinstance(y, str) else y))
        df['金额'] = pd.to_numeric(df['金额'], errors='coerce')

        global doc
        doc = Document()
        add_text_to_doc("支付宝交易分析报告", font_size=16, bold=True)
        add_text_to_doc("-" * 80)
        doc.add_paragraph()

        doc.add_heading('目录', level=1)
        doc.add_paragraph('1. 敏感词消费记录\t\t\t\t页 1')
        doc.add_paragraph('2. 行业消费记录\t\t\t\t页 2')
        doc.add_paragraph('3. 区域消费金额统计\t\t\t\t页 3')
        doc.add_paragraph('4. 交易金额前5的交易对方\t\t\t页 4')
        doc.add_paragraph('5. 交易数量前5的交易对方\t\t\t页 5')
        doc.add_page_break()

        filtered_df = df[df["交易对方"].str.contains('|'.join(sensitive_words), na=False)].copy()
        columns_to_drop = ['交易订单号', '商家订单号']
        for col in columns_to_drop:
            if col in filtered_df.columns:
                filtered_df = filtered_df.drop(columns=[col])

        if not filtered_df.empty:
            sorted_df = filtered_df.dropna(subset=["金额"]).sort_values(by="金额", ascending=False)
            sorted_df.reset_index(drop=True, inplace=True)
            doc.add_heading('1. 敏感词消费记录', level=1)
            add_text_to_doc("敏感词消费记录（按金额降序排列）：", font_size=12, bold=True)
            table = doc.add_table(rows=1, cols=len(sorted_df.columns))
            hdr_cells = table.rows[0].cells
            for i, col in enumerate(sorted_df.columns):
                hdr_cells[i].text = col
            for _, row in sorted_df.head(5).iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row):
                    row_cells[i].text = str(value)
        else:
            doc.add_heading('1. 敏感词消费记录', level=1)
            add_text_to_doc("没有敏感词消费记录")

        doc.add_page_break()

        filtered_df = df[df["交易对方"].str.contains('|'.join(industry_keywords), na=False, case=False)].copy()
        for col in columns_to_drop:
            if col in filtered_df.columns:
                filtered_df = filtered_df.drop(columns=[col])

        if not filtered_df.empty:
            sorted_df = filtered_df.dropna(subset=["金额"]).sort_values(by="金额", ascending=False)
            sorted_df.reset_index(drop=True, inplace=True)
            doc.add_heading('2. 行业消费记录', level=1)
            add_text_to_doc("行业消费记录（按金额降序排列）：", font_size=12, bold=True)
            table = doc.add_table(rows=1, cols=len(sorted_df.columns))
            hdr_cells = table.rows[0].cells
            for i, col in enumerate(sorted_df.columns):
                hdr_cells[i].text = col
            for _, row in sorted_df.head(5).iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row):
                    row_cells[i].text = str(value)
        else:
            doc.add_heading('2. 行业消费记录', level=1)
            add_text_to_doc("没有行业消费记录")

        doc.add_page_break()

        location_keywords = ["潢川县", "深圳市", "郑州市", "温县"]
        column_variants = ['交易对方', '对方', '收款方', '商户名称']
        target_column = next((col for col in column_variants if col in df.columns), None)

        if target_column is None:
            available_cols = ", ".join(df.columns.tolist())
            raise KeyError(f"找不到交易对方列，可用的列有：{available_cols}")

        filtered_df = df[df[target_column].str.contains('|'.join(location_keywords), na=False, case=False, regex=True)].copy()
        amount_variants = ['金额', '交易金额']
        amount_column = next((col for col in amount_variants if col in df.columns), None)

        if amount_column is None:
            available_cols = ", ".join(df.columns.tolist())
            raise KeyError(f"找不到金额列，可用的列有：{available_cols}")

        sorted_df = filtered_df.dropna(subset=[amount_column]).sort_values(by=amount_column, ascending=False)
        pd.set_option('display.max_rows', None)
        pd.set_option('display.unicode.ambiguous_as_wide', True)
        pd.set_option('display.unicode.east_asian_width', True)

        sorted_df['区域'] = sorted_df[target_column].apply(lambda x: extract_location(x, location_keywords))
        region_stats = sorted_df.groupby('区域')[amount_column].sum().sort_values(ascending=False).head(5)

        doc.add_heading('3. 区域消费金额统计', level=1)
        add_text_to_doc("区域消费金额统计：", font_size=12, bold=True)
        table = doc.add_table(rows=1, cols=2)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '区域'
        hdr_cells[1].text = '金额'
        for region, amount in region_stats.items():
            row_cells = table.add_row().cells
            row_cells[0].text = region
            row_cells[1].text = f"{amount:.2f}"

        plt.figure(figsize=(10, 6))
        plt.bar(region_stats.index, region_stats.values)
        plt.xlabel('区域')
        plt.ylabel('金额')
        plt.title('区域消费金额统计')
        chart_path = os.path.join(app.config['CHART_FOLDER'], 'region_chart.png')
        plt.savefig(chart_path)
        plt.close()
        doc.add_picture(chart_path, width=Inches(6))

        doc.add_page_break()

        top_5_by_amount = df.groupby('交易对方')['金额'].sum().abs().sort_values(ascending=False).head(5)
        top_5_by_count = df.groupby('交易对方')['金额'].count().sort_values(ascending=False).head(5)

        doc.add_heading('4. 交易金额前5的交易对方', level=1)
        add_text_to_doc('交易金额前5的交易对方：', font_size=12, bold=True)
        table = doc.add_table(rows=1, cols=2)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '交易对方'
        hdr_cells[1].text = '金额'
        for entity, amount in top_5_by_amount.items():
            row_cells = table.add_row().cells
            row_cells[0].text = entity
            row_cells[1].text = f"{amount:.2f}"

        plt.figure(figsize=(10, 6))
        plt.bar(top_5_by_amount.index, top_5_by_amount.values)
        plt.xlabel('交易对方')
        plt.ylabel('金额')
        plt.title('交易金额前5的交易对方')
        chart_path = os.path.join(app.config['CHART_FOLDER'], 'top_5_amount_chart.png')
        plt.savefig(chart_path)
        plt.close()
        doc.add_picture(chart_path, width=Inches(6))

        doc.add_page_break()

        doc.add_heading('5. 交易数量前5的交易对方', level=1)
        add_text_to_doc('交易数量前5的交易对方：', font_size=12, bold=True)
        table = doc.add_table(rows=1, cols=2)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '交易对方'
        hdr_cells[1].text = '交易数量'
        for entity, count in top_5_by_count.items():
            row_cells = table.add_row().cells
            row_cells[0].text = entity
            row_cells[1].text = str(count)

        plt.figure(figsize=(10, 6))
        plt.bar(top_5_by_count.index, top_5_by_count.values)
        plt.xlabel('交易对方')
        plt.ylabel('交易数量')
        plt.title('交易数量前5的交易对方')
        chart_path = os.path.join(app.config['CHART_FOLDER'], 'top_5_count_chart.png')
        plt.savefig(chart_path)
        plt.close()
        doc.add_picture(chart_path, width=Inches(6))

        doc.save(output_doc_path)

        if not os.path.exists(output_doc_path):
            raise Exception(f"DOCX保存失败，文件不存在: {output_doc_path}")

    except Exception as e:
        if os.path.exists(output_doc_path):
            try:
                os.remove(output_doc_path)
            except:
                pass
        raise


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        pdf_file = request.files['pdf_file']
        sensitive_words = request.form.get('sensitive_words', '').split(',')
        industry_keywords = request.form.get('industry_keywords', '').split(',')

        sensitive_words = [word.strip() for word in sensitive_words if word.strip()]
        industry_keywords = [word.strip() for word in industry_keywords if word.strip()]

        if not sensitive_words:
            sensitive_words = DEFAULT_SENSITIVE_WORDS
        if not industry_keywords:
            industry_keywords = DEFAULT_INDUSTRY_KEYWORDS

        if pdf_file:
            filename = secure_filename(pdf_file.filename)
            input_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            pdf_file.save(input_pdf_path)

            with pdfplumber.open(input_pdf_path) as pdf:
                first_page_text = pdf.pages[0].extract_text()

            if '微信' in first_page_text:
                output_filename = f"微信支付交易明细报告.docx"
                output_doc_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
                process_func = process_wechat_pdf
            elif '支付宝' in first_page_text:
                output_filename = f"支付宝交易分析报告.docx"
                output_doc_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
                process_func = process_alipay_pdf
            else:
                error_msg = "未在文件第一页中找到'微信'或'支付宝'关键词，请上传正确的文件。"
                return render_template('upload.html', error=error_msg,
                                       default_sensitive_words=','.join(DEFAULT_SENSITIVE_WORDS),
                                       default_industry_keywords=','.join(DEFAULT_INDUSTRY_KEYWORDS),
                                       session=session)

            try:
                os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
                process_func(input_pdf_path, output_doc_path, sensitive_words, industry_keywords)

                if os.path.exists(output_doc_path):
                    session['doc_filename'] = output_filename
                    return redirect(url_for('upload_file'))
                else:
                    raise Exception("DOCX文件未成功生成，但没有抛出具体错误")

            except Exception as e:
                error_msg = f"处理文件时出错: {str(e)}"
                print(f"错误详情: {str(e)}")
                return render_template('upload.html', error=error_msg,
                                       default_sensitive_words=','.join(DEFAULT_SENSITIVE_WORDS),
                                       default_industry_keywords=','.join(DEFAULT_INDUSTRY_KEYWORDS),
                                       session=session)

    return render_template('upload.html',
                           default_sensitive_words=','.join(DEFAULT_SENSITIVE_WORDS),
                           default_industry_keywords=','.join(DEFAULT_INDUSTRY_KEYWORDS),
                           session=session)


@app.route('/download')
def download_file():
    if 'doc_filename' not in session:
        return redirect(url_for('upload_file'))

    doc_filename = session['doc_filename']
    doc_path = os.path.join(app.config['OUTPUT_FOLDER'], doc_filename)

    if os.path.exists(doc_path):
        from flask import send_file
        return send_file(doc_path, as_attachment=True)
    else:
        session.pop('doc_filename', None)
        return render_template('upload.html', error="生成的DOCX文件不存在，请重新上传和分析",
                               default_sensitive_words=','.join(DEFAULT_SENSITIVE_WORDS),
                               default_industry_keywords=','.join(DEFAULT_INDUSTRY_KEYWORDS),
                               session=session)


if __name__ == '__main__':
    app.run(debug=True)
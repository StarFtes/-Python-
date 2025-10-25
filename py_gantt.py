import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment
import time
from functools import wraps
import os


# 计时装饰器
def timer_decorator(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        print(f"{func.__name__} 执行时间: {end_time - start_time:.4f} 秒")
        return result

    return wrapper


# 异常处理装饰器
def exception_handler_decorator(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            print(f"在 {func.__name__} 中发生错误: {str(e)}")
            raise

    return wrapper


# 数据验证装饰器 - 验证输入数据是否包含必要列
def validate_gantt_data(required_columns=None):
    if required_columns is None:
        required_columns = ['Task', 'Start', 'Duration']

    def decorator(func):
        @wraps(func)
        def wrapper(df, *args, **kwargs):
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"数据中缺少必要列: {missing_columns}")
            return func(df, *args, **kwargs)

        return wrapper

    return decorator


# 文件存在检查装饰器
def check_file_exists(func):
    @wraps(func)
    def wrapper(file_path, *args, **kwargs):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        return func(file_path, *args, **kwargs)

    return wrapper


# 读取数据函数
@exception_handler_decorator
@check_file_exists
def read_data_file(file_path, file_type=None):
    """读取数据文件，支持多种格式"""
    if file_type is None:
        # 根据文件扩展名判断类型
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            return pd.read_csv(file_path)
        elif ext in ['.xlsx', '.xls']:
            return pd.read_excel(file_path)
        else:
            raise ValueError(f"不支持的文件类型: {ext}")
    else:
        # 根据指定类型读取
        if file_type.lower() == 'csv':
            return pd.read_csv(file_path)
        elif file_type.lower() in ['excel', 'xlsx', 'xls']:
            return pd.read_excel(file_path)
        else:
            raise ValueError(f"不支持的文件类型: {file_type}")


# 数据处理函数
@exception_handler_decorator
@validate_gantt_data(['Task', 'Start', 'Duration'])
def process_gantt_data(df, start_col='Start', duration_col='Duration', task_col='Task'):
    """处理甘特图数据，确保日期格式正确"""
    # 创建数据副本以避免修改原始数据
    processed_df = df.copy()

    # 确保日期列是datetime类型
    processed_df[start_col] = pd.to_datetime(processed_df[start_col])

    # 计算结束日期
    if 'End' not in processed_df.columns:
        processed_df['End'] = processed_df[start_col] + pd.to_timedelta(processed_df[duration_col], unit='D')

    # 重命名列以符合标准格式
    column_mapping = {task_col: 'Task', start_col: 'Start', duration_col: 'Duration'}
    processed_df.rename(columns=column_mapping, inplace=True)

    # 只保留需要的列
    keep_columns = ['Task', 'Start', 'Duration', 'End']
    processed_df = processed_df[[col for col in keep_columns if col in processed_df.columns]]

    return processed_df


# 创建甘特图函数
@timer_decorator
@exception_handler_decorator
def create_gantt_chart(df, output_path='gantt_chart.png', colors=None):
    """创建甘特图"""
    # 设置中文字体支持
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial']
    plt.rcParams['axes.unicode_minus'] = False

    # 创建图形和坐标轴
    fig, ax = plt.subplots(figsize=(12, max(6, len(df) * 0.6)))  # 根据任务数量调整高度

    # 设置颜色
    if colors is None:
        colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFBE0B', '#5AB946', '#FF8E42', '#7B68EE', '#20B2AA']

    # 绘制水平条形图（甘特图）
    for i, (index, row) in enumerate(df.iterrows()):
        ax.barh(row['Task'],
                row['Duration'],
                left=row['Start'],
                color=colors[i % len(colors)],
                edgecolor='black',
                alpha=0.8)

    # 设置日期格式
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))

    # 自动调整日期刻度标签
    fig.autofmt_xdate()

    # 设置标签和标题
    ax.set_xlabel('时间')
    ax.set_ylabel('任务')
    ax.set_title('项目甘特图')

    # 添加网格线
    ax.grid(axis='x', linestyle='--', alpha=0.7)

    # 调整布局并保存
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()  # 关闭图形以避免内存泄漏

    return output_path


# 8. 导出到Excel函数
@timer_decorator
@exception_handler_decorator
def export_to_excel(df, chart_path, output_path='project_gantt.xlsx', source_file=None):
    """导出数据和图表到Excel"""
    # 创建一个新的Excel工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "项目甘特图"

    # 添加标题
    ws['A1'] = '项目甘特图'
    ws['A1'].font = Font(size=16, bold=True)
    ws.merge_cells('A1:D1')
    ws['A1'].alignment = Alignment(horizontal='center')

    # 如果有源文件信息，添加说明
    if source_file:
        ws['A2'] = f'数据来源: {source_file}'
        ws.merge_cells('A2:D2')

    # 写入数据表头
    headers = ['任务名称', '开始时间', '结束时间', '持续时间(天)']
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = Font(bold=True)

    # 写入数据
    for row_idx, (index, row) in enumerate(df.iterrows(), start=5):
        ws.cell(row=row_idx, column=1, value=row['Task'])
        ws.cell(row=row_idx, column=2, value=row['Start'])
        ws.cell(row=row_idx, column=3, value=row['End'])
        ws.cell(row=row_idx, column=4, value=row['Duration'])

    # 设置日期格式
    for row in range(5, 5 + len(df)):
        ws.cell(row=row, column=2).number_format = 'YYYY-MM-DD'
        ws.cell(row=row, column=3).number_format = 'YYYY-MM-DD'

    # 插入甘特图图片
    img = Image(chart_path)
    # 根据数据行数调整图片位置
    img_cell = f'F{4 + len(df) // 10}'  # 动态调整图片位置
    ws.add_image(img, img_cell)

    # 自动调整列宽
    for column in ['A', 'B', 'C', 'D']:
        ws.column_dimensions[column].width = 15

    # 保存Excel文件
    wb.save(output_path)
    print(f"甘特图和数据已保存到 {output_path}")
    return output_path


def user_input():
    print("欢迎使用甘特图生成工具!")
    print("请输入数据文件相对路径，以开始生成甘特图:" + "至少包含Task, Start, Duration三列" + "\n")
    file_path = input()
    return file_path


# 主函数 - 完整的甘特图创建流程
@timer_decorator
def create_gantt_from_file(file_path, output_excel='project_gantt.xlsx',
                           start_col='Start', duration_col='Duration', task_col='Task'):
    """从文件创建甘特图并导出到Excel"""
    try:
        print(f"正在处理文件: {file_path}")

        # 1. 读取数据
        df = read_data_file(file_path)
        print(f"成功读取数据，共 {len(df)} 行")

        # 2. 处理数据
        processed_df = process_gantt_data(df, start_col, duration_col, task_col)
        print("数据预处理完成")

        # 3. 创建甘特图
        chart_path = create_gantt_chart(processed_df)
        print("甘特图创建完成")

        # 4. 导出到Excel
        excel_path = export_to_excel(processed_df, chart_path, output_excel, file_path)

        print("甘特图创建和导出成功完成!")
        return excel_path

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        return None


# 示例使用
if __name__ == "__main__":
    excel_file = user_input()  # 替换为你的Excel文件路径
    result = create_gantt_from_file(excel_file, "project_gantt_from_excel.xlsx")

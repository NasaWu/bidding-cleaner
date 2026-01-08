import pandas as pd
import numpy as np


####testcase: General表处理####
# file_path = '比价导出（示例）.xlsx'
# sheet_name = 'YLC-通用'
def transform_general_table(file_path,sheet_name):
    print("开始读取文件...")
    try:
        # 读取数据，保留两层表头
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=[0, 1])
    except Exception as e:
        return f"读取Excel失败: {e}"

    # --- 步骤 1: 清洗表头 (去除空格/换行) ---
    new_columns = []
    for col in df.columns:
        c0 = str(col[0]).strip()
        c1 = str(col[1]).strip()
        new_columns.append((c0, c1))
    df.columns = pd.MultiIndex.from_tuples(new_columns)

    # --- 步骤 2: 区分 供应商列 和 基础信息列 ---
    # 逻辑：包含“份额比例”或“不含税单价”的组视为供应商
    supplier_groups = []
    top_levels = df.columns.levels[0].unique()

    for group in top_levels:
        sub_cols = df[group].columns.tolist()
        # 特征字段识别供应商
        if '份额比例' in sub_cols or '不含税单价' in sub_cols or '含税单价' in sub_cols:
            if 'Unnamed' not in group and '品项信息' not in group: 
                supplier_groups.append(group)
    
    print(f"识别到的供应商列表: {supplier_groups}")

    # --- 辅助函数: 全局搜索基础字段 ---
    def get_common_val(row_idx, target_col_names):
        if isinstance(target_col_names, str):
            target_col_names = [target_col_names]
            
        for col in df.columns:
            group_name = col[0]
            col_name = col[1]
            
            if group_name not in supplier_groups:
                if col_name in target_col_names:
                    return df.iloc[row_idx][col]
        return None

    # --- 辅助函数: 获取供应商特定字段 ---
    def get_supplier_val(row_idx, supplier, col_name):
        try:
            val = df.iloc[row_idx][(supplier, col_name)]
            return val
        except KeyError:
            return None

    new_rows = []

    # --- 步骤 3: 遍历数据行 ---
    for idx in range(len(df)):
        
        # === A. 提取基础信息 (Base Fields) ===
        base_data = {}
        
        # 通用类特有字段映射
        base_data['品项编码'] = get_common_val(idx, ['品项编码'])
        base_data['品项名称'] = get_common_val(idx, ['品项名称'])
        base_data['规格型号'] = get_common_val(idx, ['规格型号']) # 新增
        base_data['计价单位'] = get_common_val(idx, ['计价单位'])
        base_data['LC'] = get_common_val(idx, ['LC', '物流中心', '物流组']) # 主要是 LC
        base_data['需求数量'] = get_common_val(idx, ['需求数量'])
        
        base_data['是否租仓类'] = get_common_val(idx, ['是否租仓类']) # 新增
        base_data['仓库/办公室面积'] = get_common_val(idx, ['仓库/办公室面积', '面积']) # 新增，如果源表没有则为空
        
        base_data['行备注'] = get_common_val(idx, ['行备注'])
        base_data['是否附件报价'] = get_common_val(idx, ['是否附件报价'])
        base_data['批次名称'] = get_common_val(idx, ['批次名称'])

        # 日期字段 (注意：General表通常叫 价格有效期从/至，都在Base信息里)
        base_data['价格有效期从'] = get_common_val(idx, ['价格有效期从', '价格有效期', '价格有效期自'])
        base_data['价格有效期至'] = get_common_val(idx, ['价格有效期至'])

        # === B. 遍历供应商 (Supplier Fields) ===
        for supplier in supplier_groups:
            row_data = base_data.copy()
            
            # 提取供应商特有数据
            share_ratio = get_supplier_val(idx, supplier, '份额比例')
            share_qty = get_supplier_val(idx, supplier, '份额数量')
            
            row_data['供应商名称'] = supplier
            row_data['供应商编码'] = "" # 如有字典可做映射
            
            row_data['序列号'] = get_supplier_val(idx, supplier, '序列号')
            row_data['授标至'] = get_supplier_val(idx, supplier, '授标至')
            row_data['税率 (%)'] = get_supplier_val(idx, supplier, '税率')
            row_data['不含税单价'] = get_supplier_val(idx, supplier, '不含税单价')
            row_data['含税单价'] = get_supplier_val(idx, supplier, '含税单价')
            
            row_data['份额比例'] = share_ratio
            row_data['份额数量'] = share_qty
            
            # === C. 逻辑计算 ===
            # 预授标逻辑
            def is_valid_val(v):
                if pd.isna(v): return False
                s = str(v).strip()
                return s != ''
            
            if is_valid_val(share_ratio) or is_valid_val(share_qty):
                row_data['预授标'] = '是'
            else:
                row_data['预授标'] = '否'
            
            new_rows.append(row_data)

    # --- 步骤 4: 生成结果并排序 ---
    result_df = pd.DataFrame(new_rows)

    # 目标列顺序 (严格按照 General 截图)
    target_columns = [
        '序列号', '预授标', '授标至', '品项编码', '品项名称', '供应商编码', '供应商名称',
        '份额比例', '份额数量', '计价单位', '规格型号', 'LC', '税率 (%)', '不含税单价', 
        '含税单价', '价格有效期从', '价格有效期至', '需求数量', '行备注', 
        '是否租仓类', '仓库/办公室面积', '是否附件报价', '批次名称'
    ]
    
    # 确保只保留存在的列
    final_cols = [c for c in target_columns if c in result_df.columns]
    result_df = result_df[final_cols]
    
    # 日期格式优化
    for date_col in ['价格有效期从', '价格有效期至']:
        if date_col in result_df.columns:
            result_df[date_col] = pd.to_datetime(result_df[date_col], errors='coerce').dt.date

    return result_df
    # result_df.to_excel(output_path, index=False)
    # print(f"General表处理完成！文件已保存: {output_path}")

# 使用方式
# transform_general_table('general_source.xlsx', 'general_output.xlsx')
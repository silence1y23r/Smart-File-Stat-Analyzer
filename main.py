import os
import json
from collections import defaultdict
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import datetime

def count_files_in_folders():
    """通用统计：统计每个文件夹中的子文件夹和文件数量"""
    results = {}
    subfolder_set = set()  # 收集所有出现过的子文件夹名称
    other_file_types_set = set()  # 收集所有出现过的其他文件类型
    file_details = defaultdict(lambda: defaultdict(list))  # 存储文件详细信息
    other_file_details = defaultdict(lambda: defaultdict(list))  # 存储其他文件详细信息
    
    current_dir = os.getcwd()
    print("开始统计文件数量...")
    
    # 定义文件类型分类
    file_type_categories = {
        '压缩包': ['.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz'],
        'Word文档': ['.doc', '.docx', '.dot', '.dotx'],
        'Excel文档': ['.xls', '.xlsx', '.xlsm', '.xlsb', '.csv'],
        'PPT文档': ['.ppt', '.pptx', '.pps', '.ppsx'],
        'PDF文档': ['.pdf'],
        '图片': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.svg', '.webp'],
        '视频': ['.mp4', '.avi', '.mov', '.wmv', '.flv', '.mkv', '.webm'],
        '音频': ['.mp3', '.wav', '.aac', '.flac', '.ogg', '.wma'],
        '文本': ['.txt', '.md', '.log', '.ini', '.cfg', '.json', '.xml', '.yaml', '.yml'],
        '代码': ['.py', '.java', '.cpp', '.c', '.h', '.js', '.html', '.css', '.php', '.rb', '.go', '.rs'],
        '可执行文件': ['.exe', '.msi', '.dmg', '.app', '.bat', '.sh'],
        '配置文件': ['.ini', '.conf', '.cfg', '.properties'],
        '其他': []  # 其他未分类文件
    }
    
    # 创建扩展名到类型的映射
    ext_to_type = {}
    for file_type, extensions in file_type_categories.items():
        for ext in extensions:
            ext_to_type[ext.lower()] = file_type
    
    # 遍历当前目录下的所有文件夹
    for main_folder in os.listdir(current_dir):
        main_folder_path = os.path.join(current_dir, main_folder)
        
        if os.path.isdir(main_folder_path) and not main_folder.startswith('.'):
            print(f"正在处理主文件夹: {main_folder}")
            results[main_folder] = {}
            
            try:
                # 1. 首先统计子文件夹中的文件
                for item in os.listdir(main_folder_path):
                    item_path = os.path.join(main_folder_path, item)
                    
                    if os.path.isdir(item_path):
                        subfolder_set.add(item)
                        file_count = 0
                        file_list = []
                        
                        # 统计子文件夹中的文件
                        for file_item in os.listdir(item_path):
                            file_item_path = os.path.join(item_path, file_item)
                            if os.path.isfile(file_item_path):
                                file_count += 1
                                file_list.append(file_item)
                        
                        results[main_folder][item] = file_count
                        file_details[main_folder][item] = file_list
                
                # 2. 统计主文件夹下的其他文件（非文件夹）
                other_files_count = 0
                other_files_by_type = defaultdict(int)
                other_files_list_by_type = defaultdict(list)
                
                for item in os.listdir(main_folder_path):
                    item_path = os.path.join(main_folder_path, item)
                    
                    if os.path.isfile(item_path):
                        other_files_count += 1
                        
                        # 获取文件扩展名
                        _, ext = os.path.splitext(item)
                        ext = ext.lower()
                        
                        # 确定文件类型
                        if ext in ext_to_type:
                            file_type = ext_to_type[ext]
                        else:
                            # 检查是否是已知类型但不在扩展名列表中
                            for type_name, extensions in file_type_categories.items():
                                if ext in extensions:
                                    file_type = type_name
                                    break
                            else:
                                file_type = '其他'
                        
                        # 添加到类型集合
                        type_key = f"其他-{file_type}"
                        other_file_types_set.add(type_key)
                        
                        # 统计数量
                        other_files_by_type[type_key] += 1
                        
                        # 记录文件名
                        other_files_list_by_type[type_key].append(item)
                
                # 将其他文件的统计结果添加到结果中
                for type_key, count in other_files_by_type.items():
                    results[main_folder][type_key] = count
                
                # 保存其他文件的详细信息
                other_file_details[main_folder] = other_files_list_by_type
                
                # 3. 统计主文件夹下的总文件数（包括子文件夹中的文件和其他文件）
                total_subfolder_files = sum(results[main_folder].get(sf, 0) for sf in subfolder_set)
                total_other_files = sum(results[main_folder].get(ot, 0) for ot in other_file_types_set)
                total_files = total_subfolder_files + total_other_files
                results[main_folder]['文件总数'] = total_files
                
            except Exception as e:
                print(f"  错误: 无法访问 {main_folder}: {e}")
                results[main_folder] = {"错误": f"访问失败: {e}"}
    
    # 按字母顺序排序子文件夹和其他文件类型
    subfolders = sorted(list(subfolder_set))
    other_file_types = sorted(list(other_file_types_set))
    
    # 合并所有列
    all_columns = subfolders + other_file_types
    
    return results, all_columns, file_details, other_file_details

def create_excel_report(results, all_columns, file_details, other_file_details):
    """创建Excel报告"""
    # 获取当前时间用于文件名
    current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"文件夹统计报告_{current_time}.xlsx"
    
    # 创建Excel写入器
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        # ==================== 工作表1: 文件夹统计 ====================
        detailed_data = []
        
        for main_folder, data in results.items():
            if "错误" in data:
                # 如果有错误，特殊处理
                row = {"主文件夹": main_folder}
                for col in all_columns:
                    row[col] = "访问错误"
                row["文件总数"] = "N/A"
                detailed_data.append(row)
                continue
            
            row = {"主文件夹": main_folder}
            row_total = 0
            
            for col in all_columns:
                count = data.get(col, 0)
                row[col] = count
                if isinstance(count, int):
                    row_total += count
            
            row["文件总数"] = data.get('文件总数', row_total)
            detailed_data.append(row)
        
        # 创建DataFrame并排序
        df_detailed = pd.DataFrame(detailed_data)
        df_detailed = df_detailed.sort_values(by="主文件夹")
        
        # 添加汇总行
        if detailed_data:
            total_row = {"主文件夹": "总计"}
            
            for col in all_columns:
                # 计算该列在所有主文件夹中的总数
                total_count = sum(
                    0 if not isinstance(row.get(col, 0), int) else row.get(col, 0)
                    for row in detailed_data if "主文件夹" in row and row["主文件夹"] != "总计"
                )
                total_row[col] = total_count
            
            # 计算总文件数
            total_files = sum(
                0 if not isinstance(row.get("文件总数", 0), (int, float)) else row.get("文件总数", 0)
                for row in detailed_data if "主文件夹" in row and row["主文件夹"] != "总计"
            )
            total_row["文件总数"] = total_files
            
            # 将汇总行添加到DataFrame
            df_detailed = pd.concat([df_detailed, pd.DataFrame([total_row])], ignore_index=True)
        
        # 写入Excel
        df_detailed.to_excel(writer, sheet_name='文件夹统计', index=False)
        
        # ==================== 工作表2: 文件清单 ====================
        file_list_data = []
        
        # 2.1 添加子文件夹中的文件
        for main_folder, subfolders_dict in file_details.items():
            for subfolder, files in subfolders_dict.items():
                if files:
                    # 将文件列表转换为字符串，用换行符分隔
                    files_str = "\n".join(files)
                    file_list_data.append({
                        "主文件夹": main_folder,
                        "位置类型": "子文件夹",
                        "位置/类型": subfolder,
                        "文件数量": len(files),
                        "文件列表": files_str
                    })
                else:
                    file_list_data.append({
                        "主文件夹": main_folder,
                        "位置类型": "子文件夹",
                        "位置/类型": subfolder,
                        "文件数量": 0,
                        "文件列表": "(空文件夹)"
                    })
        
        # 2.2 添加主文件夹下的其他文件
        for main_folder, other_files_dict in other_file_details.items():
            for file_type, files in other_files_dict.items():
                if files:
                    # 去除"其他-"前缀
                    display_type = file_type.replace("其他-", "")
                    files_str = "\n".join(files)
                    file_list_data.append({
                        "主文件夹": main_folder,
                        "位置类型": "其他文件",
                        "位置/类型": display_type,
                        "文件数量": len(files),
                        "文件列表": files_str
                    })
        
        if file_list_data:
            df_file_list = pd.DataFrame(file_list_data)
            df_file_list = df_file_list.sort_values(by=["主文件夹", "位置类型", "位置/类型"])
            df_file_list.to_excel(writer, sheet_name='文件清单', index=False)
        else:
            # 如果没有文件，创建一个空的工作表
            df_file_list = pd.DataFrame({"说明": ["未找到任何文件"]})
            df_file_list.to_excel(writer, sheet_name='文件清单', index=False)
        
        # ==================== 工作表3: 汇总分析 ====================
        if all_columns:
            summary_data = []
            
            # 分离子文件夹和其他文件类型
            subfolders = [col for col in all_columns if not col.startswith("其他-")]
            other_types = [col for col in all_columns if col.startswith("其他-")]
            
            # 3.1 子文件夹的汇总
            print("\n正在生成子文件夹汇总...")
            for subfolder in subfolders:
                # 统计有多少个主文件夹包含这个子文件夹
                main_folders_with_subfolder = sum(
                    1 for data in results.values() 
                    if isinstance(data, dict) and subfolder in data and data[subfolder] > 0
                )
                
                # 统计该子文件夹中的总文件数
                total_files_in_subfolder = sum(
                    data.get(subfolder, 0) for data in results.values() 
                    if isinstance(data, dict) and isinstance(data.get(subfolder, 0), int)
                )
                
                # 计算平均文件数
                avg_files = total_files_in_subfolder / main_folders_with_subfolder if main_folders_with_subfolder > 0 else 0
                
                # 计算文件类型分布
                file_extensions = defaultdict(int)
                for main_folder, subfolders_dict in file_details.items():
                    if subfolder in subfolders_dict:
                        for file in subfolders_dict[subfolder]:
                            ext = os.path.splitext(file)[1].lower() or "(无扩展名)"
                            file_extensions[ext] += 1
                
                # 取最多的3种文件类型
                common_exts = sorted(file_extensions.items(), key=lambda x: x[1], reverse=True)[:3]
                common_exts_str = "; ".join([f"{ext}:{count}" for ext, count in common_exts])
                
                summary_data.append({
                    "分类": "子文件夹",
                    "名称": subfolder,
                    "总文件数": total_files_in_subfolder,
                    "包含的主文件夹数": main_folders_with_subfolder,
                    "平均文件数": round(avg_files, 2),
                    "常见文件类型": common_exts_str,
                    "说明": f"子文件夹: {subfolder}"
                })
            
            # 3.2 其他文件类型的汇总
            print("正在生成其他文件类型汇总...")
            for other_type in other_types:
                display_name = other_type.replace("其他-", "")
                
                # 统计有多少个主文件夹包含这个文件类型
                main_folders_with_type = sum(
                    1 for data in results.values() 
                    if isinstance(data, dict) and other_type in data and data[other_type] > 0
                )
                
                # 统计该文件类型的总文件数
                total_files_of_type = sum(
                    data.get(other_type, 0) for data in results.values() 
                    if isinstance(data, dict) and isinstance(data.get(other_type, 0), int)
                )
                
                # 计算平均文件数
                avg_files = total_files_of_type / main_folders_with_type if main_folders_with_type > 0 else 0
                
                # 收集所有文件名
                all_files = []
                for main_folder, other_files_dict in other_file_details.items():
                    if other_type in other_files_dict:
                        all_files.extend(other_files_dict[other_type])
                
                # 统计文件扩展名分布
                file_extensions = defaultdict(int)
                for file in all_files:
                    ext = os.path.splitext(file)[1].lower() or "(无扩展名)"
                    file_extensions[ext] += 1
                
                # 取最多的3种文件扩展名
                common_exts = sorted(file_extensions.items(), key=lambda x: x[1], reverse=True)[:3]
                common_exts_str = "; ".join([f"{ext}:{count}" for ext, count in common_exts])
                
                summary_data.append({
                    "分类": "其他文件",
                    "名称": display_name,
                    "总文件数": total_files_of_type,
                    "包含的主文件夹数": main_folders_with_type,
                    "平均文件数": round(avg_files, 2),
                    "常见文件类型": common_exts_str,
                    "说明": f"主文件夹下的{display_name}"
                })
            
            # 添加总计行
            if summary_data:
                # 子文件夹总计
                subfolder_total = sum(item["总文件数"] for item in summary_data if item["分类"] == "子文件夹")
                subfolder_count = len([item for item in summary_data if item["分类"] == "子文件夹"])
                
                # 其他文件总计
                other_total = sum(item["总文件数"] for item in summary_data if item["分类"] == "其他文件")
                other_count = len([item for item in summary_data if item["分类"] == "其他文件"])
                
                # 整体总计
                total_files_all = subfolder_total + other_total
                total_main_folders = len(results)
                
                summary_data.append({
                    "分类": "总计",
                    "名称": "全部",
                    "总文件数": total_files_all,
                    "包含的主文件夹数": total_main_folders,
                    "平均文件数": round(total_files_all / total_main_folders, 2) if total_main_folders > 0 else 0,
                    "常见文件类型": f"子文件夹: {subfolder_count}种, 其他文件: {other_count}种",
                    "说明": f"总计: {total_files_all}个文件"
                })
            
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='汇总分析', index=False)
        else:
            df_summary = pd.DataFrame({"说明": ["未找到任何子文件夹或文件"]})
            df_summary.to_excel(writer, sheet_name='汇总分析', index=False)
        
        # 获取工作簿和工作表对象用于样式设置
        workbook = writer.book
        sheet_detailed = workbook['文件夹统计']
        sheet_file_list = workbook['文件清单']
        sheet_summary = workbook['汇总分析']
        
        # ==================== 设置样式 ====================
        # 定义样式
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        
        # 设置详细统计表的样式
        print("正在设置详细统计表样式...")
        # 设置列宽
        for column in sheet_detailed.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            # 根据列内容调整宽度
            if column[0].value == "主文件夹":
                adjusted_width = min(max_length + 2, 25)
            elif column[0].value == "文件总数":
                adjusted_width = 12
            else:
                adjusted_width = min(max_length + 2, 20)
            
            sheet_detailed.column_dimensions[column_letter].width = adjusted_width
        
        # 设置表头样式
        for cell in sheet_detailed[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = border
        
        # 设置数据样式
        for row in sheet_detailed.iter_rows(min_row=2, max_row=sheet_detailed.max_row):
            for cell in row:
                cell.alignment = align_center
                cell.border = border
                
                # 为汇总行添加特殊样式
                if cell.row == sheet_detailed.max_row and cell.column == 1:
                    cell.font = Font(bold=True, color="FF0000")
                elif cell.row == sheet_detailed.max_row:
                    cell.font = Font(bold=True)
                
                # 为访问错误添加特殊样式
                if cell.value == "访问错误":
                    cell.font = Font(color="FF0000")
                    cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                # 为0值添加特殊样式
                if cell.value == 0 and isinstance(cell.value, int):
                    cell.font = Font(color="808080")
                
                # 为"其他-"开头的列添加特殊背景色
                if cell.column > 1 and sheet_detailed.cell(row=1, column=cell.column).value and sheet_detailed.cell(row=1, column=cell.column).value.startswith("其他-"):
                    if cell.row == 1:
                        # 表头使用不同的颜色
                        cell.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
                    elif cell.value and cell.value != 0:
                        # 数据行使用浅紫色背景
                        cell.fill = PatternFill(start_color="E6E0EC", end_color="E6E0EC", fill_type="solid")
        
        # 设置文件清单表的样式
        print("正在设置文件清单表样式...")
        if file_list_data:
            # 设置列宽
            sheet_file_list.column_dimensions['A'].width = 20  # 主文件夹
            sheet_file_list.column_dimensions['B'].width = 12  # 位置类型
            sheet_file_list.column_dimensions['C'].width = 20  # 位置/类型
            sheet_file_list.column_dimensions['D'].width = 12  # 文件数量
            sheet_file_list.column_dimensions['E'].width = 50  # 文件列表
            
            # 设置表头样式
            for cell in sheet_file_list[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = align_center
                cell.border = border
            
            # 设置数据样式
            for row in sheet_file_list.iter_rows(min_row=2, max_row=sheet_file_list.max_row):
                for cell in row:
                    cell.border = border
                    
                    # 根据列设置对齐方式
                    if cell.column <= 4:  # 前4列居中对齐
                        cell.alignment = align_center
                    else:  # 第5列左对齐并允许换行
                        cell.alignment = align_left
                    
                    # 为0值添加特殊样式
                    if cell.value == 0 and isinstance(cell.value, int):
                        cell.font = Font(color="808080")
                    
                    # 为空文件夹添加特殊样式
                    if cell.value == "(空文件夹)":
                        cell.font = Font(color="808080", italic=True)
                    
                    # 为不同位置类型添加不同背景色
                    if cell.column == 2:  # 位置类型列
                        if cell.value == "子文件夹":
                            cell.fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
                        elif cell.value == "其他文件":
                            cell.fill = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
        
        # 设置汇总分析表的样式
        print("正在设置汇总分析表样式...")
        # 设置列宽
        for column in sheet_summary.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 30)
            sheet_summary.column_dimensions[column_letter].width = adjusted_width
        
        # 设置表头样式
        for cell in sheet_summary[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = border
        
        # 设置数据样式
        for row in sheet_summary.iter_rows(min_row=2, max_row=sheet_summary.max_row):
            for cell in row:
                cell.alignment = align_center
                cell.border = border
                
                # 为汇总行添加特殊样式
                if cell.row == sheet_summary.max_row and cell.column == 1:
                    cell.font = Font(bold=True, color="FF0000")
                elif cell.row == sheet_summary.max_row:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                
                # 根据分类添加不同背景色
                if cell.column == 1 and cell.value:  # 分类列
                    if cell.value == "子文件夹":
                        cell.fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
                    elif cell.value == "其他文件":
                        cell.fill = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
                    elif cell.value == "总计":
                        cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    print(f"\nExcel文件已生成: {excel_filename}")
    return excel_filename

def create_simple_excel(results, all_columns, file_details, other_file_details):
    """创建简化的Excel文件（如果openpyxl有问题）"""
    try:
        return create_excel_report(results, all_columns, file_details, other_file_details)
    except Exception as e:
        print(f"创建美化Excel时出错: {e}")
        print("正在尝试创建简化版Excel...")
        
        # 创建简化版
        current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        simple_filename = f"文件夹统计报告_简化版_{current_time}.xlsx"
        
        # 创建详细统计表
        detailed_data = []
        for main_folder, data in results.items():
            if "错误" in data:
                row = {"主文件夹": main_folder}
                for col in all_columns:
                    row[col] = "访问错误"
                row["文件总数"] = "N/A"
                detailed_data.append(row)
                continue
            
            row = {"主文件夹": main_folder}
            row_total = 0
            for col in all_columns:
                count = data.get(col, 0)
                row[col] = count
                if isinstance(count, int):
                    row_total += count
            row["文件总数"] = data.get('文件总数', row_total)
            detailed_data.append(row)
        
        df_detailed = pd.DataFrame(detailed_data)
        
        # 创建文件清单表
        file_list_data = []
        
        # 添加子文件夹中的文件
        for main_folder, subfolders_dict in file_details.items():
            for subfolder, files in subfolders_dict.items():
                if files:
                    files_str = "\n".join(files)
                    file_list_data.append({
                        "主文件夹": main_folder,
                        "位置类型": "子文件夹",
                        "位置/类型": subfolder,
                        "文件数量": len(files),
                        "文件列表": files_str
                    })
        
        # 添加主文件夹下的其他文件
        for main_folder, other_files_dict in other_file_details.items():
            for file_type, files in other_files_dict.items():
                if files:
                    display_type = file_type.replace("其他-", "")
                    files_str = "\n".join(files)
                    file_list_data.append({
                        "主文件夹": main_folder,
                        "位置类型": "其他文件",
                        "位置/类型": display_type,
                        "文件数量": len(files),
                        "文件列表": files_str
                    })
        
        df_file_list = pd.DataFrame(file_list_data)
        
        # 保存到Excel
        with pd.ExcelWriter(simple_filename, engine='openpyxl') as writer:
            df_detailed.to_excel(writer, sheet_name='文件夹统计', index=False)
            df_file_list.to_excel(writer, sheet_name='文件清单', index=False)
        
        print(f"简化版Excel文件已生成: {simple_filename}")
        return simple_filename

def main():
    print("=" * 60)
    print("通用文件夹统计工具 - Excel版")
    print("=" * 60)
    print("说明：")
    print("1. 请将此脚本放在包含所有主文件夹的目录中")
    print("2. 每个主文件夹下的子文件夹和其他文件都将被统计")
    print("3. 其他文件将按类型分类（压缩包、Word文档、Excel文档等）")
    print("4. 生成的Excel包含3个工作表:")
    print("   - 文件夹统计: 主文件夹的详细文件数量统计")
    print("   - 文件清单: 每个子文件夹和文件类型的具体文件名")
    print("   - 汇总分析: 统计摘要和分析")
    print("5. 正在统计...\n")
    
    try:
        # 安装必要的库（如果用户同意）
        try:
            import pandas as pd
            from openpyxl import Workbook
        except ImportError:
            print("检测到缺少必要的库，正在尝试安装...")
            import subprocess
            import sys
            
            # 安装pandas和openpyxl
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
            print("安装完成！")
            
            # 重新导入
            import pandas as pd
            from openpyxl import Workbook
        
        # 执行统计
        results, all_columns, file_details, other_file_details = count_files_in_folders()
        
        if not results:
            print("\n未找到任何主文件夹！")
            print("请确保脚本放在正确的目录中。")
            return
        
        # 分离子文件夹和其他文件类型
        subfolders = [col for col in all_columns if not col.startswith("其他-")]
        other_types = [col for col in all_columns if col.startswith("其他-")]
        
        # 创建Excel报告
        excel_file = create_simple_excel(results, all_columns, file_details, other_file_details)
        
        # 显示统计摘要
        print("\n" + "=" * 60)
        print("统计完成！")
        print("=" * 60)
        print(f"已处理主文件夹数: {len(results)}")
        print(f"发现的子文件夹类型数: {len(subfolders)}")
        print(f"发现的其他文件类型数: {len(other_types)}")
        print(f"生成的Excel文件: {excel_file}")
        
        # 计算总文件数
        total_all_files = 0
        for main_folder, data in results.items():
            if isinstance(data, dict):
                total = data.get('文件总数', 0)
                if isinstance(total, (int, float)):
                    total_all_files += total
        
        # 显示其他文件类型统计
        if other_types:
            print(f"\n其他文件类型统计:")
            for other_type in other_types[:10]:  # 只显示前10个
                type_display = other_type.replace("其他-", "")
                type_total = sum(
                    data.get(other_type, 0) for data in results.values() 
                    if isinstance(data, dict) and isinstance(data.get(other_type, 0), int)
                )
                print(f"  {type_display}: {type_total} 个文件")
            
            if len(other_types) > 10:
                print(f"  ... 还有 {len(other_types) - 10} 种其他文件类型")
        
        print(f"\n总文件数统计: {total_all_files} 个文件")
        
    except Exception as e:
        print(f"发生错误: {e}")
        print("\n请尝试以下解决方案:")
        print("1. 安装必要库: pip install pandas openpyxl")
        print("2. 确保脚本放在包含主文件夹的目录中")
        print("3. 确保有足够的权限访问这些文件夹")

if __name__ == "__main__":
    main()
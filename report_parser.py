import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Dict, List, Tuple, Set
import threading

# 尝试导入PDF处理库
try:
    import pdfplumber

    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import PyPDF2

    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False

try:
    from fontTools.ttLib import TTCollection

    HAS_FONTTOOLS = True
except ImportError:
    HAS_FONTTOOLS = False


class TaggedConcentration(float):
    """带状态的浓度值：对 GUI 的阈值比较/格式化零影响（float 子类），同时携带 status 供无GUI调用读取。

    status ∈ {'检出', '未检出', '未校正'}。原始解析路径把 N.D./未校正 统一塌缩成 0.0，
    丢失了状态；数据采集的「项目别名」规则需要区分它们（未检出→报出 blank 值）。
    """

    def __new__(cls, value, status='检出', raw=None):
        obj = float.__new__(cls, value)
        obj.status = status
        obj.raw = raw  # 原始数值串(保留报告有效位/末尾0，如0.100)，仅检出时有值
        return obj


class PDFTextExtractor:
    """PDF文本提取器"""

    def extract_text_from_pdf(self, file_path: str) -> str:
        """从PDF文件中提取文本"""
        text = ""

        # 优先使用pdfplumber，因为它通常能提取更多文本
        if HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        # 尝试使用不同的提取策略
                        page_text = page.extract_text()
                        if not page_text or len(page_text.strip()) < 10:
                            # 如果文本提取不理想，尝试使用布局模式
                            page_text = page.extract_text(layout=True)
                        if page_text:
                            text += page_text + "\n"
                if text.strip():
                    return text
            except Exception as e:
                print(f"使用pdfplumber提取文本失败: {e}")

        # 回退到PyPDF2
        if HAS_PYPDF2:
            try:
                with open(file_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    for page in pdf_reader.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                return text
            except Exception as e:
                print(f"使用PyPDF2提取文本失败: {e}")

        # 如果两种方法都失败
        raise Exception("无法从PDF中提取文本，请确保已安装PyPDF2或pdfplumber")

    def extract_tables_from_pdf(self, file_path: str) -> List:
        """从PDF中提取表格数据"""
        tables = []
        if HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        page_tables = page.extract_tables()
                        if page_tables:
                            tables.extend(page_tables)
            except Exception as e:
                print(f"提取PDF表格时出错: {e}")
        return tables


class ChemicalReportAnalyzer:
    def __init__(self):
        self._init_parser()
        self.root = tk.Tk()
        self.root.title("报告解析器")
        # 进一步减小界面大小
        self.root.geometry("700x500")
        self.setup_ui()

    def _init_parser(self):
        """初始化解析相关状态（不依赖Tk），供无GUI子类复用全部解析方法。"""
        # 检查PDF支持状态
        self.has_pdf_support = HAS_PDFPLUMBER or HAS_PYPDF2

        # 标记物
        self.markers = set()

        self.file_contents = {}

        # 初始化PDF提取器
        if self.has_pdf_support:
            self.pdf_extractor = PDFTextExtractor()
        else:
            self.pdf_extractor = None

        # 改进浓度匹配模式
        self.concentration_pattern = re.compile(r'(\d+\.\d+)\s*(ppm|mg/L|μg/mL|ng/ml)', re.IGNORECASE)
        self.concentration_pattern_int = re.compile(r'(\d+)\s*(ppm|mg/L|μg/mL|ng/ml)', re.IGNORECASE)
        # ponytail: 前后向否定断言保证 ND/N.D. 是独立标记，否则 "I{nd}eno"、"E{nd}osulfan" 等化合物名会被误判为未检出
        self.nd_pattern = re.compile(r'(?<![A-Za-z])N\.?\s*D\.?(?![A-Za-z])', re.IGNORECASE)
        self.uncorrected_pattern = re.compile(r'未校正', re.IGNORECASE)

        # 有效的化合物名称模式 - 使用更宽松的正则表达式
        self.valid_compound_pattern = re.compile(r'^[A-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+$')

    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="4")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择区域 - 更紧凑的设计
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="4")
        file_frame.pack(fill=tk.X, pady=(0, 4))

        # 文件选择按钮和开始分析按钮在同一行
        file_top_frame = ttk.Frame(file_frame)
        file_top_frame.pack(fill=tk.X, pady=1)

        ttk.Button(file_top_frame, text="选择报告文件",
                   command=self.select_files).pack(side=tk.LEFT, padx=(0, 8))

        self.file_count_label = ttk.Label(file_top_frame, text="未选择文件")
        self.file_count_label.pack(side=tk.LEFT, padx=(0, 10))

        # 开始分析按钮放在同一行右对齐，使用相同格式
        ttk.Button(file_top_frame, text="开始分析",
                   command=self.start_analysis).pack(side=tk.RIGHT)

        # 分析设置区域 - 更紧凑的设计
        settings_frame = ttk.LabelFrame(main_frame, text="分析设置", padding="4")
        settings_frame.pack(fill=tk.X, pady=(0, 4))

        # 阈值设置和标记物设置在同一行
        settings_top_frame = ttk.Frame(settings_frame)
        settings_top_frame.pack(fill=tk.X, pady=1)

        # 阈值设置
        threshold_frame = ttk.Frame(settings_top_frame)
        threshold_frame.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(threshold_frame, text="检测阈值:").pack(side=tk.LEFT)
        self.threshold_var = tk.StringVar(value="0.1")
        ttk.Entry(threshold_frame, textvariable=self.threshold_var, width=6).pack(side=tk.LEFT, padx=3)
        ttk.Label(threshold_frame, text="(≥此值为阳性)").pack(side=tk.LEFT)

        # 标记物设置
        marker_frame = ttk.Frame(settings_top_frame)
        marker_frame.pack(side=tk.LEFT)

        self.marker_var = tk.BooleanVar()
        marker_check = ttk.Checkbutton(marker_frame, text="排除标记物",
                                       variable=self.marker_var, command=self.toggle_marker_input)
        marker_check.pack(side=tk.LEFT)

        self.marker_entry = ttk.Entry(marker_frame, width=18, state="disabled")
        self.marker_entry.pack(side=tk.LEFT, padx=3)
        self.marker_entry.bind("<Return>", self.update_markers)

        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="分析结果", padding="4")
        result_frame.pack(fill=tk.BOTH, expand=True)

        # 创建标签页
        notebook = ttk.Notebook(result_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        # 摘要标签页 - 减小高度
        summary_tab = ttk.Frame(notebook, padding="4")
        notebook.add(summary_tab, text="摘要")

        self.summary_text = scrolledtext.ScrolledText(summary_tab, height=8, wrap=tk.WORD)
        self.summary_text.pack(fill=tk.BOTH, expand=True)

        # 详细结果标签页 - 减小高度
        detail_tab = ttk.Frame(notebook, padding="4")
        notebook.add(detail_tab, text="详细结果")

        self.detail_text = scrolledtext.ScrolledText(detail_tab, height=8, wrap=tk.WORD)
        self.detail_text.pack(fill=tk.BOTH, expand=True)

        # 化合物列表标签页 - 减小高度
        compound_tab = ttk.Frame(notebook, padding="4")
        notebook.add(compound_tab, text="所有化合物")

        self.compound_text = scrolledtext.ScrolledText(compound_tab, height=8, wrap=tk.WORD)
        self.compound_text.pack(fill=tk.BOTH, expand=True)

        # 调试信息标签页 - 减小高度
        debug_tab = ttk.Frame(notebook, padding="4")
        notebook.add(debug_tab, text="调试信息")

        self.debug_text = scrolledtext.ScrolledText(debug_tab, height=8, wrap=tk.WORD)
        self.debug_text.pack(fill=tk.BOTH, expand=True)

        # 状态栏 - 移除边框阴影
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief="flat")
        status_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(2, 0))

    def toggle_marker_input(self):
        """切换标记物输入框状态"""
        if self.marker_var.get():
            self.marker_entry.config(state="normal")
            self.status_var.set("请输入标记物名称，用逗号分隔")
        else:
            self.marker_entry.config(state="disabled")
            self.markers.clear()
            self.status_var.set("标记物排除已禁用")

    def update_markers(self, event=None):
        """更新标记物列表"""
        marker_text = self.marker_entry.get().strip()
        if marker_text:
            self.markers = {marker.strip() for marker in marker_text.split(",")}
            self.status_var.set(f"已设置 {len(self.markers)} 个标记物")
        else:
            self.markers.clear()
            self.status_var.set("标记物列表已清空")

    def select_files(self):
        """选择报告文件"""
        filetypes = [("文本文件", "*.txt"), ("所有文件", "*.*")]
        if self.has_pdf_support:
            filetypes.insert(0, ("PDF文件", "*.pdf"))

        files = filedialog.askopenfilenames(
            title="选择报告文件",
            filetypes=filetypes
        )

        if files:
            self.file_contents.clear()
            self.status_var.set(f"正在读取 {len(files)} 个文件...")

            # 在新线程中读取文件
            threading.Thread(target=self.read_files, args=(files,), daemon=True).start()

    def read_files(self, files):
        """读取文件内容"""
        successful_files = 0

        for file_path in files:
            try:
                # 检查文件扩展名
                _, ext = os.path.splitext(file_path)

                if ext.lower() == '.pdf':
                    # 如果没有PDF支持，跳过PDF文件
                    if not self.has_pdf_support:
                        continue

                    # 使用PDF提取器处理PDF文件
                    content = self.pdf_extractor.extract_text_from_pdf(file_path)
                else:
                    # 处理文本文件
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()

                if content:
                    filename = os.path.basename(file_path)
                    self.file_contents[filename] = {
                        'content': content,
                        'path': file_path
                    }
                    successful_files += 1

            except Exception as e:
                print(f"读取文件 {file_path} 时出错: {e}")

        # 更新UI
        self.root.after(0, self.update_file_count, successful_files, len(files))

    def update_file_count(self, successful, total):
        """更新文件计数显示"""
        self.file_count_label.config(text=f"已选择 {successful}/{total} 个文件")
        if successful < total and not self.has_pdf_support:
            self.status_var.set(f"成功读取 {successful} 个文件 (跳过PDF文件)")
        else:
            self.status_var.set(f"成功读取 {successful} 个文件")

    def start_analysis(self):
        """开始分析"""
        if not self.file_contents:
            messagebox.showwarning("警告", "请先选择要分析的文件")
            return

        try:
            threshold = float(self.threshold_var.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的阈值数字")
            return

        # 更新标记物列表
        if self.marker_var.get():
            self.update_markers()

        self.status_var.set("正在分析报告...")

        # 在新线程中进行分析
        threading.Thread(target=self.analyze_reports_thread, args=(threshold,), daemon=True).start()

    def analyze_reports_thread(self, threshold):
        """在新线程中分析报告"""
        try:
            detected, not_detected, details, debug_info = self.analyze_reports(self.file_contents, threshold)
            self.root.after(0, self.display_results, detected, not_detected, details, debug_info, threshold)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("分析错误", str(e)))

    def analyze_reports(self, file_contents: Dict[str, dict], threshold: float) -> Tuple[
        List[str], List[str], Dict, str]:
        """分析所有报告"""
        detected_reports = []
        not_detected_reports = []
        detailed_results = {}
        debug_info = "=== 调试信息 ===\n\n"

        for filename, file_info in file_contents.items():
            content = file_info['content']
            file_path = file_info['path']
            compounds, file_debug_info = self.parse_report_content(content, filename, file_path)
            debug_info += file_debug_info

            # 过滤掉标记物
            target_compounds = {}
            for name, conc in compounds.items():
                if not self.is_marker(name):
                    target_compounds[name] = conc

            # 检查是否有超过等于阈值的化合物
            detected_compounds = {name: conc for name, conc in target_compounds.items()
                                  if conc >= threshold}  # 改为大于等于

            if detected_compounds:
                detected_reports.append(filename)
            else:
                not_detected_reports.append(filename)

            detailed_results[filename] = {
                'all_compounds': compounds,
                'target_compounds': target_compounds,
                'detected_compounds': detected_compounds,
                'has_detection': len(detected_compounds) > 0
            }

        return detected_reports, not_detected_reports, detailed_results, debug_info

    def parse_report_content(self, content: str, filename: str, file_path: str) -> Tuple[Dict[str, float], str]:
        """解析报告内容，提取化合物和浓度信息"""
        compounds = {}
        debug_info = f"文件: {filename}\n"

        # 首先尝试MH的TSCA报告解析方法
        mh_compounds, mh_debug = self.parse_tsca_format(content, filename, file_path)
        debug_info += mh_debug

        if mh_compounds:
            compounds.update(mh_compounds)
            debug_info += f"  MH TSCA格式解析成功: {len(mh_compounds)} 个化合物\n"

        # 如果MH方法没有找到化合物，尝试原有的多种解析策略
        if not compounds:
            lines = content.split('\n')
            strategies = [
                self.parse_compact_format,  # 先尝试紧凑格式（PAE报告）
                self.parse_standard_quant_report,  # 再尝试标准定量报告格式
                self.parse_table_format,  # 再尝试表格格式
                self.parse_gcfid_format,  # GC-FID(cid 乱码)简短定量报告
                self.parse_standard_format,  # 最后尝试标准格式
            ]

            for strategy in strategies:
                result, strategy_debug = strategy(lines, filename, file_path)
                debug_info += strategy_debug
                if result:
                    # 只添加新解析的化合物，避免覆盖
                    for name, conc in result.items():
                        if name not in compounds and self.is_valid_compound(name):
                            compounds[name] = conc
                    debug_info += f"  策略 {strategy.__name__} 成功解析了 {len(result)} 个化合物\n"

        debug_info += f"  最终解析结果: {len(compounds)} 个化合物\n\n"
        return compounds, debug_info

    def parse_tsca_format(self, content: str, filename: str, file_path: str) -> Tuple[Dict[str, float], str]:
        """解析TSCA报告格式（MH的解析方法）"""
        compounds = {}
        debug_info = f"  尝试MH TSCA格式解析...\n"

        # 提取化合物和浓度信息
        tsca_compounds, istd_info = self.extract_compounds_and_concentrations(content, file_path)
        debug_info += f"    提取到 {len(tsca_compounds)} 个化合物\n"

        # 转换为标准格式
        for compound_name, concentration_str in tsca_compounds.items():
            if self.is_valid_compound(compound_name):
                # 解析浓度值
                if concentration_str == "N.D.":
                    compounds[compound_name] = TaggedConcentration(0.0, '未检出')
                else:
                    # 提取数值部分
                    conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', concentration_str, re.IGNORECASE)
                    if conc_match:
                        try:
                            concentration = float(conc_match.group(1))
                            compounds[compound_name] = TaggedConcentration(concentration, '检出', conc_match.group(1))
                        except ValueError:
                            compounds[compound_name] = TaggedConcentration(0.0, '未检出')
                    else:
                        compounds[compound_name] = TaggedConcentration(0.0, '未检出')

        return compounds, debug_info

    def parse_gcfid_format(self, lines, filename, file_path):
        """解析 GC-FID 简短定量报告(BXW 苯系物 / TDI 等)。
        列固定：名称/保留时间/峰面积/响应因子/含量[mg/L]。SimSun 未嵌入→pdfplumber 返回 (cid:NNNN)，
        用 _decode_cid_text(系统 simsun.ttc cmap)动态解出中文化合物名，零硬编码。
        内标行响应因子恒=1.000(定义)，据此动态剔除。浓度取末列「含量」；含量0→未检出。
        不依赖表头解码——部分重导出报告表头 cid 被重编为乱码，故按列结构(名称+≥3数值)定位。"""
        compounds = {}
        debug_info = "  尝试 GC-FID 格式解析...\n"
        content = '\n'.join(lines) if isinstance(lines, list) else (lines or '')
        # ponytail: 仅 cid 乱码报告触发(正常报告中文可直抽，避免无谓跑表格)
        if content.count('(cid:') < 3:
            return {}, debug_info + "    非 cid 乱码，跳过\n"

        def _is_num(s):
            try:
                float(s)
                return True
            except (ValueError, TypeError):
                return False

        tables = self.pdf_extractor.extract_tables_from_pdf(file_path)
        for table in tables:
            for row in (table or []):
                cells = [_decode_cid_text(str(c)) if c else '' for c in row]
                if len(cells) < 4:
                    continue
                name = cells[0].strip()
                if not name or name == '名称' or not self.is_valid_compound(name):
                    continue
                # 数据行：名称后跟 ≥3 个数值(RT/峰面积/响应因子/含量)
                if len([c for c in cells[1:] if _is_num(c)]) < 3:
                    continue
                resp, conc = cells[-2], cells[-1]   # 响应因子(倒数第二)/含量(末列)
                # 内标：响应因子==1.000(内标相对自身的定义)；动态剔除，不靠写死名单
                try:
                    if abs(float(resp) - 1.0) < 1e-6:
                        debug_info += f"    跳过内标: {name}\n"
                        continue
                except ValueError:
                    pass
                try:
                    val = float(conc)
                except ValueError:
                    continue
                if val > 0:
                    compounds[name] = TaggedConcentration(val, '检出', conc)
                else:
                    compounds[name] = TaggedConcentration(0.0, '未检出')
        if compounds:
            debug_info += f"  GC-FID 格式解析成功: {len(compounds)} 个化合物\n"
        return compounds, debug_info

    def extract_compounds_and_concentrations(self, content: str, file_path: str) -> Tuple[Dict[str, str], Dict]:
        """从TSCA报告中提取化合物和浓度信息"""
        compounds = {}
        istd_info = {}

        # 首先尝试从表格中提取数据
        tables = self.pdf_extractor.extract_tables_from_pdf(file_path)
        table_compounds = self.extract_from_tables(tables, content)
        compounds.update(table_compounds)

        # 如果表格提取失败，尝试从文本中提取
        if not compounds:
            text_compounds, istd_info = self.extract_from_text(content)
            compounds.update(text_compounds)

        # 对于外标法报告，使用专门的提取方法
        if '外标法' in content:
            external_compounds = self.extract_from_external_standard(content)
            compounds.update(external_compounds)

        # 清理无效的化合物名称
        compounds = self.clean_compounds(compounds)

        return compounds, istd_info

    def extract_from_tables(self, tables: List, content: str) -> Dict[str, str]:
        """从PDF表格中提取化合物和浓度"""
        compounds = {}

        for table in tables:
            if not table:
                continue

            # 查找表头行
            header_row = -1
            for i, row in enumerate(table):
                if row and any(
                        cell and any(keyword in str(cell) for keyword in ['名称', '化合物', 'ISTD', '浓度']) for cell in
                        row):
                    header_row = i
                    break

            if header_row == -1:
                continue

            # 确定名称和浓度所在的列
            name_col = -1
            conc_col = -1

            for j, cell in enumerate(table[header_row]):
                if cell and any(keyword in str(cell) for keyword in ['名称', '化合物']):
                    name_col = j
                elif cell and any(keyword in str(cell) for keyword in ['浓度', '最终浓度']):
                    conc_col = j

            # 如果没找到浓度列，尝试最后一列
            if conc_col == -1 and len(table[header_row]) > 1:
                conc_col = len(table[header_row]) - 1

            # 提取数据行
            for i in range(header_row + 1, len(table)):
                row = table[i]
                if not row or len(row) <= max(name_col, conc_col):
                    continue

                name_cell = row[name_col] if name_col < len(row) else None
                conc_cell = row[conc_col] if conc_col < len(row) else None

                if name_cell and self.is_valid_compound(name_cell):
                    # 处理浓度值
                    concentration = "N.D."
                    if conc_cell:
                        if 'N.D.' in str(conc_cell):
                            concentration = "N.D."
                        else:
                            conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', str(conc_cell), re.IGNORECASE)
                            if conc_match:
                                concentration = f"{conc_match.group(1)} {conc_match.group(2)}"

                    compounds[name_cell] = concentration

        return compounds

    def extract_from_text(self, content: str) -> Tuple[Dict[str, str], Dict]:
        """从文本中提取化合物和浓度"""
        compounds = {}
        istd_info = {}
        lines = content.split('\n')

        # 查找表格区域
        table_start = -1
        for i, line in enumerate(lines):
            # 内标法表头
            if '名称' in line and 'ISTD' in line and '最终浓度单位' in line:
                table_start = i
                break
            # 外标法表头
            elif '名称' in line and '浓度' in line:
                table_start = i
                break

        if table_start == -1:
            return compounds, istd_info

        # 提取内标信息（仅对内标法）
        if 'ISTD' in content:
            for i in range(max(0, table_start - 5), min(table_start + 5, len(lines))):
                line = lines[i]
                if 'ISTD浓度单位' in line:
                    # 提取内标浓度
                    conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', line, re.IGNORECASE)
                    if conc_match:
                        istd_info['concentration'] = f"{conc_match.group(1)} {conc_match.group(2)}"

        # 提取化合物数据
        for i in range(table_start + 1, min(table_start + 50, len(lines))):
            line = lines[i].strip()
            if not line or '生成时间' in line or '样品色谱图' in line:
                break

            # 尝试解析化合物行
            # 内标法格式: 化合物名称 ISTD名称 ISTD保留时间 ISTD离子对 ISTD响应 ISTD浓度单位 保留时间 离子对 响应 最终浓度
            # 外标法格式: 化合物名称 保留时间 离子对 响应 浓度
            parts = re.split(r'\s+', line)
            if len(parts) >= 3:
                compound = parts[0]

                # 查找浓度值（通常在最后一列或倒数第二列）
                concentration = None
                for part in parts[-3:]:  # 检查最后三列
                    if 'N.D.' in part:
                        concentration = "N.D."
                        break
                    conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', part, re.IGNORECASE)
                    if conc_match:
                        concentration = f"{conc_match.group(1)} {conc_match.group(2)}"
                        break
                # ponytail: 必须真正命中浓度(N.D.或单位)，否则跳过——避免把数据表后的色谱图
                # 坐标轴行(stnuoC/223.0,/100/x104)误当化合物(它们没有浓度标记)。
                if concentration is None:
                    continue
                if self.is_valid_compound(compound):
                    compounds[compound] = concentration

        return compounds, istd_info

    def extract_from_external_standard(self, content: str) -> Dict[str, str]:
        """专门处理外标法报告的化合物提取"""
        compounds = {}
        lines = content.split('\n')

        # 查找外标法表格区域
        table_start = -1
        for i, line in enumerate(lines):
            if '名称' in line and 'RT' in line and '响应' in line and '离子对' in line and '最终浓度单位' in line:
                table_start = i
                break

        if table_start == -1:
            # 如果没有找到完整表头，尝试简化表头
            for i, line in enumerate(lines):
                if '名称' in line and '最终浓度单位' in line:
                    table_start = i
                    break

        if table_start == -1:
            return compounds

        # 从表格区域提取化合物
        for i in range(table_start + 1, min(table_start + 30, len(lines))):
            line = lines[i].strip()
            if not line or any(end_marker in line for end_marker in ['生成时间', '样品色谱图', '=====']):
                break

            # 尝试分割行数据
            parts = re.split(r'\s{2,}', line)  # 按多个空格分割
            if len(parts) < 5:  # 如果分割后列数不够，尝试按单个空格分割
                parts = line.split()

            if len(parts) >= 5:
                # 第一列是化合物名称，最后一列是浓度
                compound = parts[0]
                concentration = parts[-1]

                # 处理特殊情况：如果化合物名称包含逗号，可能需要合并
                if ',' in compound and len(parts) > 5:
                    # 尝试合并前几列作为化合物名称
                    compound = ' '.join(parts[0:2])
                    concentration = parts[-1]

                # 处理中文化合物名称
                if self.is_valid_compound(compound):
                    # 标准化浓度格式
                    if 'N.D.' in concentration:
                        compounds[compound] = "N.D."
                    else:
                        conc_match = re.search(r'([\d.]+)\s*(mg/L|μg/mL|ng/ml)', concentration, re.IGNORECASE)
                        if conc_match:
                            compounds[compound] = f"{conc_match.group(1)} {conc_match.group(2)}"

        return compounds

    def clean_compounds(self, compounds: Dict[str, str]) -> Dict[str, str]:
        """清理化合物字典，移除无效的化合物名称"""
        cleaned = {}
        for compound, concentration in compounds.items():
            # 排除反向文本和其他无效名称
            if self.is_valid_compound(compound) and not self.is_excluded_compound(compound):
                cleaned[compound] = concentration
        return cleaned

    def is_excluded_compound(self, text: str) -> bool:
        """检查是否是应该排除的化合物名称"""
        # 排除列表 - 包含常见无效名称和反向文本
        exclude_list = [
            'ISTD', 'RT', '响应', '离子对', '面积', 'Counts', 'min',
            '样品', '分析', '报告', '名称', '浓度', '最终浓度', '数据', '文件',
            '采集方法', '样品瓶', '其他', '稀释', '体积', '类型', '批处理',
            'stnuoC', 'niM', 'x10', 'x10^', 'TIC', 'EIC', 'Scan',
            '比值', '生成时间', '样品色谱图', 'Selected', 'Ion', 'SIM'
        ]

        # 检查是否是反向文本
        reversed_text = text[::-1]
        if any(exclude in reversed_text for exclude in ['Counts', 'Min', 'RT']):
            return True

        if any(exclude in text for exclude in exclude_list):
            return True

        return False

    # 原有的解析方法保持不变，只添加file_path参数
    def parse_compact_format(self, lines, filename: str, file_path: str = None):
        """解析紧凑格式的报告（如PAE报告）"""
        compounds = {}
        debug_info = f"  尝试紧凑格式解析...\n"

        # 查找目标化合物区域
        in_target_section = False
        compound_count = 0
        parsing_finished = False

        for i, line in enumerate(lines):
            if parsing_finished:
                break

            line = line.strip()

            # 检测目标化合物区域
            if '目标化合物' in line or 'Target Compounds' in line:
                in_target_section = True
                debug_info += f"    第 {i + 1} 行: 进入目标化合物区域\n"
                continue

            # 检查报告解析结束标志
            if '----------' in line or '==========' in line or '定性峰超出范围' in line:
                if in_target_section:
                    debug_info += f"    第 {i + 1} 行: 遇到结束标志，停止解析\n"
                    parsing_finished = True
                    break
                continue

            # 跳过空行
            if not line:
                continue

            # 跳过注释行和其他非化合物行
            if self.should_skip_line(line):
                debug_info += f"    第 {i + 1} 行: 跳过非化合物行\n"
                continue

            # 只在目标化合物区域解析化合物行
            if in_target_section:
                # 检查是否是化合物行（包含数字和可能的浓度值）
                if re.search(r'\d+\)\s+', line) and (
                        'N.D.' in line or 'mg/L' in line or '未校正' in line or re.search(r'\d+\.\d+\s+mg/L', line)):
                    compound_name, concentration, parse_debug = self.parse_compact_compound_line(line)
                    debug_info += f"    第 {i + 1} 行: {parse_debug}\n"
                    if compound_name and self.is_valid_compound(compound_name) and compound_name not in compounds:
                        compounds[compound_name] = concentration
                        compound_count += 1

        debug_info += f"    解析了 {compound_count} 个化合物\n"
        return compounds, debug_info

    def should_skip_line(self, line: str) -> bool:
        """判断是否应该跳过该行（非化合物行）"""
        skip_patterns = [
            r'\(#\)', r'定性峰', r'手动积分', r'已求和的信号',
            r'PAES-CAL', r'SVHC-CAL', r'SCCP-MCCP-CAL', r'PCN-CAL', r'Fri', r'Mon', r'Tue', r'Wed', r'Thu', r'Sat',
            r'Sun',
            r'^\d{4}-\d{2}-\d{2}', r'\d{2}:\d{2}:\d{2}', r'^\s*$', r'^---', r'^===',
            r'\.M\s+\d', r'周一', r'周二', r'周三', r'周四', r'周五', r'周六', r'周日',
            r'内标', r'ISTD', r'Internal Standard'  # 跳过内标相关行
        ]

        for pattern in skip_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                return True

        # 检查是否包含日期时间格式
        if re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', line) or re.search(r'\d{1,2}:\d{2}:\d{2}', line):
            return True

        return False

    def parse_compact_compound_line(self, line: str) -> Tuple[str, float, str]:
        """解析紧凑格式报告中的化合物行（如PAE报告）"""
        debug_info = f"解析紧凑格式行: {line[:50]}..."

        # 提取化合物名称
        compound_name = self.extract_compound_name_from_compact_line(line)
        if not compound_name:
            debug_info += " -> 无法提取化合物名称"
            return "", 0.0, debug_info

        # 使用新的浓度解析方法
        concentration, _, status, raw = self.parse_concentration(line)

        if status in ['未检出', '未校正']:
            debug_info += f" -> {status}: {compound_name}"
            return compound_name, TaggedConcentration(0.0, status), debug_info

        debug_info += f" -> {compound_name}: {concentration}"
        return compound_name, TaggedConcentration(concentration, '检出', raw), debug_info

    def extract_compound_name_from_compact_line(self, line: str) -> str:
        """从紧凑格式行中提取化合物名称"""
        # 移除行首的数字和括号
        cleaned = re.sub(r'^\d+\)\s*', '', line)  # 移除 "1) " 格式

        # 特别处理希腊字母α的情况
        # 匹配包含希腊字母α的化合物名称
        greek_pattern = r'^([αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+)'
        match = re.match(greek_pattern, cleaned)
        if match:
            compound_name = match.group(1).strip()
            # 进一步清理：确保名称合理
            if len(compound_name) >= 2 and not compound_name.replace('-', '').replace(',', '').replace('.', '').replace(
                    '\'', '').replace('(', '').replace(')', '').replace('（', '').replace('）', '').replace('[',
                                                                                                          '').replace(
                ']', '').replace('α', '').replace('β', '').replace('γ', '').replace('δ', '').replace('ε', '').replace(
                'ζ', '').replace('η', '').replace('θ', '').replace('ι', '').replace('κ', '').replace('λ', '').replace(
                'μ', '').replace('ν', '').replace('ξ', '').replace('ο', '').replace('π', '').replace('ρ', '').replace(
                'σ', '').replace('τ', '').replace('υ', '').replace('φ', '').replace('χ', '').replace('ψ', '').replace(
                'ω', '').isdigit():
                return compound_name

        # 如果上述方法失败，使用通用清理方法
        return self.clean_compound_name(cleaned)

    def parse_standard_quant_report(self, lines, filename: str, file_path: str = None):
        """解析标准定量报告格式（如TS25101323001-氯苯-检出.pdf）"""
        compounds = {}
        debug_info = f"  尝试标准定量报告解析...\n"

        # 查找目标化合物区域
        in_target_section = False
        compound_count = 0
        parsing_finished = False

        for i, line in enumerate(lines):
            if parsing_finished:
                break

            line = line.strip()

            # 检测目标化合物区域
            if '目标化合物' in line or 'Target Compounds' in line:
                in_target_section = True
                debug_info += f"    第 {i + 1} 行: 进入目标化合物区域\n"
                continue

            # 检查报告解析结束标志
            if '----------' in line or '==========' in line or '定性峰超出范围' in line:
                if in_target_section:
                    debug_info += f"    第 {i + 1} 行: 遇到结束标志，停止解析\n"
                    parsing_finished = True
                    break
                continue

            # 跳过空行
            if not line:
                continue

            # 跳过注释行和其他非化合物行
            if self.should_skip_line(line):
                debug_info += f"    第 {i + 1} 行: 跳过非化合物行\n"
                continue

            # 只在目标化合物区域解析化合物行
            if in_target_section:
                # 检查是否是化合物行（包含数字和单位）
                if re.search(r'\d+\.\d+\s+mg/L', line) or 'N.D.' in line or '0.000 mg/L' in line or re.search(
                        r'\d+\)\s+[\wα\-\[\]]', line):
                    compound_name, concentration, parse_debug = self.parse_quant_compound_line(line)
                    debug_info += f"    第 {i + 1} 行: {parse_debug}\n"
                    if compound_name and self.is_valid_compound(compound_name) and compound_name not in compounds:
                        compounds[compound_name] = concentration
                        compound_count += 1

        debug_info += f"    解析了 {compound_count} 个化合物\n"
        return compounds, debug_info

    def parse_quant_compound_line(self, line: str) -> Tuple[str, float, str]:
        """解析定量报告中的化合物行"""
        debug_info = f"解析: {line[:50]}..."

        # 提取化合物名称
        compound_name = self.extract_compound_name_from_quant_line(line)
        if not compound_name:
            debug_info += " -> 无法提取化合物名称"
            return "", 0.0, debug_info

        # 使用新的浓度解析方法
        concentration, _, status, raw = self.parse_concentration(line)

        if status in ['未检出', '未校正']:
            debug_info += f" -> {status}: {compound_name}"
            return compound_name, TaggedConcentration(0.0, status), debug_info

        debug_info += f" -> {compound_name}: {concentration}"
        return compound_name, TaggedConcentration(concentration, '检出', raw), debug_info

    def extract_compound_name_from_quant_line(self, line: str) -> str:
        """从定量报告行中提取化合物名称"""
        # 移除行首的数字和括号
        cleaned = re.sub(r'^\d+\)\s*', '', line)  # 移除 "1) " 格式

        # 特别处理希腊字母α的情况
        # 匹配包含希腊字母α的化合物名称
        greek_pattern = r'^([αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+)'
        match = re.match(greek_pattern, cleaned)
        if match:
            compound_name = match.group(1).strip()
            # 进一步清理：确保名称合理
            if len(compound_name) >= 2 and not compound_name.replace('-', '').replace(',', '').replace('.', '').replace(
                    '\'', '').replace('(', '').replace(')', '').replace('（', '').replace('）', '').replace('[',
                                                                                                          '').replace(
                ']', '').replace('α', '').replace('β', '').replace('γ', '').replace('δ', '').replace('ε', '').replace(
                'ζ', '').replace('η', '').replace('θ', '').replace('ι', '').replace('κ', '').replace('λ', '').replace(
                'μ', '').replace('ν', '').replace('ξ', '').replace('ο', '').replace('π', '').replace('ρ', '').replace(
                'σ', '').replace('τ', '').replace('υ', '').replace('φ', '').replace('χ', '').replace('ψ', '').replace(
                'ω', '').isdigit():
                return compound_name

        # 如果上述方法失败，使用通用清理方法
        return self.clean_compound_name(cleaned)

    def parse_table_format(self, lines, filename: str, file_path: str = None):
        """解析表格格式的报告（如TS25110765001A-PAHS-检出.pdf）"""
        compounds = {}
        debug_info = f"  尝试表格格式解析...\n"

        # 查找表格开始
        in_table = False
        header_found = False
        compound_count = 0
        in_target_section = False
        parsing_finished = False

        for i, line in enumerate(lines):
            if parsing_finished:
                break

            line = line.strip()

            # 检测目标化合物区域
            if '目标化合物' in line or 'Target Compounds' in line:
                in_target_section = True
                debug_info += f"    第 {i + 1} 行: 进入目标化合物区域\n"
                continue

            # 检查报告解析结束标志
            if '----------' in line or '==========' in line or '定性峰超出范围' in line:
                if in_target_section:
                    debug_info += f"    第 {i + 1} 行: 遇到结束标志，停止解析\n"
                    parsing_finished = True
                    break
                continue

            # 跳过注释行和其他非化合物行
            if self.should_skip_line(line):
                debug_info += f"    第 {i + 1} 行: 跳过非化合物行\n"
                continue

            # 检测表格开始
            if in_target_section and '| 化合物' in line and '保留时间' in line:
                in_table = True
                header_found = True
                debug_info += f"    第 {i + 1} 行: 进入表格区域\n"
                continue

            # 解析表格行 - 只在目标化合物区域解析
            if in_target_section and in_table and line and '|' in line:
                # 跳过表头行
                if header_found:
                    header_found = False
                    continue

                # 解析表格行
                compound_name, concentration, parse_debug = self.parse_table_row(line)
                debug_info += f"    第 {i + 1} 行: {parse_debug}\n"
                if compound_name and self.is_valid_compound(compound_name) and compound_name not in compounds:
                    compounds[compound_name] = concentration
                    compound_count += 1

        debug_info += f"    解析了 {compound_count} 个化合物\n"
        return compounds, debug_info

    def parse_table_row(self, line: str) -> Tuple[str, float, str]:
        """解析表格行"""
        debug_info = f"解析表格行: {line[:50]}..."

        # 提取化合物名称
        compound_name = self.extract_compound_name_from_table(line)
        if not compound_name:
            debug_info += " -> 无法提取化合物名称"
            return "", 0.0, debug_info

        # 使用新的浓度解析方法
        concentration, _, status, raw = self.parse_concentration(line)

        if status in ['未检出', '未校正']:
            debug_info += f" -> {status}: {compound_name}"
            return compound_name, TaggedConcentration(0.0, status), debug_info

        debug_info += f" -> {compound_name}: {concentration}"
        return compound_name, TaggedConcentration(concentration, '检出', raw), debug_info

    def extract_compound_name_from_table(self, line: str) -> str:
        """从表格行中提取化合物名称"""
        # 分割表格列
        parts = [part.strip() for part in line.split('|') if part.strip()]

        if len(parts) >= 1:
            # 第一列是化合物名称，可能包含数字和括号
            name_part = parts[0]
            # 使用新的化合物名称清理方法
            return self.clean_compound_name(name_part)

        return ""

    def parse_standard_format(self, lines, filename: str, file_path: str = None):
        """解析标准格式的报告"""
        compounds = {}
        debug_info = f"  尝试标准格式解析...\n"

        # 查找目标化合物区域
        in_target_section = False
        compound_count = 0
        parsing_finished = False

        for i, line in enumerate(lines):
            if parsing_finished:
                break

            line = line.strip()

            # 检测目标化合物区域
            if '目标化合物' in line or 'Target Compounds' in line:
                in_target_section = True
                debug_info += f"    第 {i + 1} 行: 进入目标化合物区域\n"
                continue

            # 检查报告解析结束标志
            if '----------' in line or '==========' in line or '定性峰超出范围' in line:
                if in_target_section:
                    debug_info += f"    第 {i + 1} 行: 遇到结束标志，停止解析\n"
                    parsing_finished = True
                    break
                continue

            # 跳过表格分隔线
            if '---' in line or '===' in line or not line:
                continue

            # 跳过注释行和其他非化合物行
            if self.should_skip_line(line):
                debug_info += f"    第 {i + 1} 行: 跳过非化合物行\n"
                continue

            # 只在目标化合物区域解析化合物行
            if in_target_section:
                compound_name, concentration, parse_debug = self.parse_standard_line(line)
                debug_info += f"    第 {i + 1} 行: {parse_debug}\n"
                if compound_name and self.is_valid_compound(compound_name) and compound_name not in compounds:
                    compounds[compound_name] = concentration
                    compound_count += 1

        debug_info += f"    解析了 {compound_count} 个化合物\n"
        return compounds, debug_info

    def parse_standard_line(self, line: str) -> Tuple[str, float, str]:
        """解析标准格式的化合物行"""
        debug_info = f"解析标准行: {line[:50]}..."

        # 提取化合物名称
        compound_name = self.extract_compound_name_from_standard_line(line)
        if not compound_name:
            debug_info += " -> 无法提取化合物名称"
            return "", 0.0, debug_info

        # 使用新的浓度解析方法
        concentration, _, status, raw = self.parse_concentration(line)

        if status in ['未检出', '未校正']:
            debug_info += f" -> {status}: {compound_name}"
            return compound_name, TaggedConcentration(0.0, status), debug_info

        debug_info += f" -> {compound_name}: {concentration}"
        return compound_name, TaggedConcentration(concentration, '检出', raw), debug_info

    def extract_compound_name_from_standard_line(self, line: str) -> str:
        """从标准格式行中提取化合物名称"""
        # 移除行首的数字和括号
        cleaned = re.sub(r'^\d+\)\s*', '', line)  # 移除 "1) " 格式

        # 特别处理希腊字母α的情况
        # 匹配包含希腊字母α的化合物名称
        greek_pattern = r'^([αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+)'
        match = re.match(greek_pattern, cleaned)
        if match:
            compound_name = match.group(1).strip()
            # 进一步清理：确保名称合理
            if len(compound_name) >= 2 and not compound_name.replace('-', '').replace(',', '').replace('.', '').replace(
                    '\'', '').replace('(', '').replace(')', '').replace('（', '').replace('）', '').replace('[',
                                                                                                          '').replace(
                ']', '').replace('α', '').replace('β', '').replace('γ', '').replace('δ', '').replace('ε', '').replace(
                'ζ', '').replace('η', '').replace('θ', '').replace('ι', '').replace('κ', '').replace('λ', '').replace(
                'μ', '').replace('ν', '').replace('ξ', '').replace('ο', '').replace('π', '').replace('ρ', '').replace(
                'σ', '').replace('τ', '').replace('υ', '').replace('φ', '').replace('χ', '').replace('ψ', '').replace(
                'ω', '').isdigit():
                return compound_name

        # 如果上述方法失败，使用通用清理方法
        return self.clean_compound_name(cleaned)

    def parse_concentration(self, concentration_str: str) -> Tuple[float, str, str, str]:
        """
        解析浓度字符串，返回浓度值、单位、状态和原始数值串(保留报告里的有效位/末尾0)
        """
        if not concentration_str or concentration_str.strip() == '':
            return 0.0, '', '未检测', None

        concentration_str = str(concentration_str).strip()

        # 检查是否为未检出
        if self.nd_pattern.search(concentration_str):
            return 0.0, '', '未检出', None

        # 检查是否为未校正
        if self.uncorrected_pattern.search(concentration_str):
            return 0.0, '', '未校正', None

        # 提取浓度值和单位
        # 首先尝试匹配小数+单位
        match = self.concentration_pattern.search(concentration_str)
        if match:
            value = float(match.group(1))
            unit = match.group(2)
            # 放宽浓度值范围限制（从0.00001到10000）
            if 0.00001 <= value <= 10000:
                return value, unit, '检出', match.group(1)

        # 特殊处理：如果文本被错误分割（如"0.3 6 mg/L"）
        parts = concentration_str.split()
        for i in range(len(parts) - 2):
            # 检查是否可能是被分割的小数
            # 模式：数字+点+空格+数字+空格+单位
            if ('.' in parts[i] and
                    parts[i].replace('.', '').replace(' ', '').isdigit() and
                    parts[i + 1].isdigit() and
                    parts[i + 2].lower() in ['mg/l', 'ppm', 'μg/ml', 'ng/ml']):
                try:
                    # 合并数字部分，去掉中间的空格
                    value_str = parts[i] + parts[i + 1]  # "0.3" + "6" = "0.36"
                    value = float(value_str)
                    unit = parts[i + 2]
                    # 放宽浓度值范围限制
                    if 0.00001 <= value <= 10000:
                        return value, unit, '检出', value_str
                except ValueError:
                    continue

        # 尝试匹配整数+单位（只有在没有小数的情况下）
        match = self.concentration_pattern_int.search(concentration_str)
        if match:
            value = float(match.group(1))
            unit = match.group(2)
            # 放宽浓度值范围限制
            if 0.00001 <= value <= 10000:
                return value, unit, '检出', match.group(1)

        return 0.0, '', '未检出', None

    def clean_compound_name(self, name: str) -> str:
        """
        清理化合物名称，移除编号和多余空格
        """
        # 移除类似 "1) "、"17) " 这样的编号
        name = re.sub(r'^\d+\)\s*', '', name)
        # 移除类似 "107) " 这样的编号
        name = re.sub(r'^\d+\)', '', name)
        # 移除保留时间等数字前缀
        name = re.sub(r'^\d+\.\d+\s+', '', name)
        # 移除多余空格
        name = re.sub(r'\s+', ' ', name).strip()

        # 进一步清理：移除末尾的点和空格
        name = re.sub(r'[.\s]+$', '', name)

        # 提取有效的化合物名称部分（包括希腊字母）
        greek_pattern = r'^([αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+)'
        match = re.match(greek_pattern, name)
        if match:
            return match.group(1).strip()

        return name

    def is_valid_compound(self, compound_name: str) -> bool:
        """检查是否为有效的化合物名称"""
        if not compound_name:
            return False
        # 单个中文字符也是合法化合物名(如 苯/酚/蒽)；仅拒绝单字节杂质
        if len(compound_name) < 2 and not re.match(r'^[一-鿿]$', compound_name):
            return False

        # 检查是否符合化合物名称模式（包括希腊字母）
        greek_pattern = r'^[αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+$'
        if re.match(greek_pattern, compound_name):
            return True

        # 检查是否为中文化合物名称
        if re.match(r'^[\u4e00-\u9fff]+$', compound_name):
            return True

        # 允许包含数字和特殊字符的化合物名称（包括希腊字母）
        extended_pattern = r'^[αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z][αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]*$'
        if re.match(extended_pattern, compound_name):
            return True

        # MH的化合物名称模式
        compound_patterns = [
            r'^[A-Z]{2,}-\d+$',  # 如 DIBP-223, DEHP-279
            r'^[A-Z]{3,}$',  # 如 MOCA, TSCA
            r'^[A-Z]{2,}\d*$',  # 如 BB, MDCA
            r'^[A-Z]+\(\d+:\d+\)$',  # 如 PIP(3:1)
            r'^[A-Z]+-\d+[A-Z]*$',  # 如 BBP-206
            r'^[\u4e00-\u9fff]+$',  # 纯中文化合物名称
            r'^[\u4e00-\u9fff][\u4e00-\u9fff\w\-\(\):,]+$',  # 中文混合，包含逗号
            r'^[\d,]+-[\u4e00-\u9fff]+$',  # 如 2,4,6-三叔丁基苯酚
            r'^[αβγδεζηθικλμνξοπρστυφχψω]+.*$',  # 以希腊字母开头的化合物
        ]

        for pattern in compound_patterns:
            if re.match(pattern, compound_name):
                return True

        return False

    def is_marker(self, compound_name: str) -> bool:
        """判断是否为标记物"""
        if not self.marker_var.get():
            return False
        return any(marker.lower() in compound_name.lower() for marker in self.markers)

    def display_results(self, detected: List[str], not_detected: List[str], details: Dict, debug_info: str,
                        threshold: float):
        """显示分析结果"""
        # 更新摘要标签页
        self.summary_text.delete(1.0, tk.END)

        summary = f"=== 分析结果 (阈值: {threshold}) ===\n\n"

        # 显示标记物排除状态
        if self.marker_var.get() and self.markers:
            summary += f"排除的标记物: {', '.join(self.markers)}\n\n"

        summary += f"✅ 阳性报告 ({len(detected)}):\n"
        for report in detected:
            detected_count = len(details[report]['detected_compounds'])
            summary += f"  ✓ {report} (检测到 {detected_count} 个化合物)\n"

        summary += f"\n❌ 阴性报告 ({len(not_detected)}):\n"
        for report in not_detected:
            summary += f"  ✗ {report}\n"

        self.summary_text.insert(1.0, summary)

        # 更新详细结果标签页
        self.detail_text.delete(1.0, tk.END)

        detail_content = f"=== 详细分析结果 ===\n\n"

        for filename, result in details.items():
            status = "✅ 阳性" if result['has_detection'] else "❌ 阴性"
            detail_content += f"📄 文件: {filename} [{status}]\n"

            if result['detected_compounds']:
                detail_content += "   检测到的化合物:\n"
                for name, conc in result['detected_compounds'].items():
                    detail_content += f"     ✓ {name}: {conc}\n"
            else:
                detail_content += "   未检测到超过阈值的化合物\n"

            detail_content += "\n"

        self.detail_text.insert(1.0, detail_content)

        # 更新化合物列表标签页
        self.compound_text.delete(1.0, tk.END)

        compound_content = "=== 所有检测到的化合物 ===\n\n"

        all_compounds = {}
        for filename, result in details.items():
            for name, conc in result['all_compounds'].items():
                if name not in all_compounds:
                    all_compounds[name] = []
                all_compounds[name].append((filename, conc))

        # 按化合物名称排序
        for compound in sorted(all_compounds.keys()):
            compound_content += f"🔬 {compound}:\n"
            for filename, conc in all_compounds[compound]:
                status = "标记物" if self.is_marker(compound) else "目标物"
                compound_content += f"   {filename}: {conc} [{status}]\n"
            compound_content += "\n"

        self.compound_text.insert(1.0, compound_content)

        # 更新调试信息标签页
        self.debug_text.delete(1.0, tk.END)
        self.debug_text.insert(1.0, debug_info)

        self.status_var.set(f"分析完成 - 阳性: {len(detected)}, 阴性: {len(not_detected)}")

    def run(self):
        """运行应用程序"""
        self.root.mainloop()


class _HeadlessReportAnalyzer(ChemicalReportAnalyzer):
    """无Tk根窗与UI——供 parse_pdf_report 在主进程内调用，复用父类全部解析方法。"""

    def __init__(self):
        self._init_parser()  # 仅初始化解析器，不创建 tk.Tk、不 setup_ui


# GC-FID 简短定量报告的 SimSun 字体未嵌入 PDF，pdfplumber 抽出 (cid:NNNN) 乱码。
# CID=GID(CIDFontType2 Identity)，用系统 simsun.ttc 的 cmap 反查 GID→Unicode 动态解出中文——
# 化合物名取自报告本身，零硬编码。需 Windows + 系统装 SimSun；否则解码为空(回退原 cid 串)。
_SIMSUN_CIDS_RE = re.compile(r'\(cid:(\d+)\)')
_SIMSUN_GID2UNI = None  # 懒加载缓存 {gid: unicode}


def _simsun_gid2uni():
    """惰性加载系统 simsun.ttc 的 GID→Unicode 映射(首次读字体文件，之后复用)。失败返回 {}。"""
    global _SIMSUN_GID2UNI
    if _SIMSUN_GID2UNI is not None:
        return _SIMSUN_GID2UNI
    _SIMSUN_GID2UNI = {}
    if not HAS_FONTTOOLS:
        return _SIMSUN_GID2UNI
    try:
        font = TTCollection(r"C:\Windows\Fonts\simsun.ttc").fonts[0]
        cmap = font.getBestCmap()               # {unicode: glyphname}
        glyph_order = font.getGlyphOrder()      # index=GID -> glyphname
        name2uni = {}
        for uni, gname in cmap.items():
            name2uni.setdefault(gname, uni)
        _SIMSUN_GID2UNI = {gid: name2uni.get(gn) for gid, gn in enumerate(glyph_order)}
    except Exception:
        _SIMSUN_GID2UNI = {}
    return _SIMSUN_GID2UNI


def _decode_cid_text(s):
    """把 pdfplumber 的 (cid:NNNN) 用系统 SimSun cmap 解成中文；非 cid 串原样返回。"""
    if not s or '(cid:' not in s:
        return s
    g2u = _simsun_gid2uni()

    def _repl(m):
        uni = g2u.get(int(m.group(1)))
        return chr(uni) if uni else ''

    return _SIMSUN_CIDS_RE.sub(_repl, s)


def _extract_column_headers(text):
    """识别报告表头行（含'化合物'+'保留时间'，如 MassHunter QT 报告），按空白拆成表头列表。
    用于「数据采集」按 equipRelativeTitle 匹配 PDF 表头定位填充列。"""
    for line in (text or '').splitlines():
        if '化合物' in line and '保留时间' in line:
            return line.split()
    return []


# 文件名稀释倍数：-NNX（如 TS24031359001A-50X → 50）。要求前置连字符，避免样品编号误匹配。
# 仅用于识别稀释报告文件 + 取倍数填「稀释」列；parse_pdf_report 不据此缩放浓度（LIMS 自算）。
_DILUTION_RE = re.compile(r'-(\d+)\s*[Xx]')


def _dilution_factor(filename):
    """从文件名提取稀释倍数，取最后一个 -NNX；无则 1.0。TS...A-50X → 50.0"""
    stem = os.path.splitext(os.path.basename(filename))[0]
    nums = _DILUTION_RE.findall(stem)
    return float(nums[-1]) if nums else 1.0


def _split_content_dilution(sid):
    """样品段 id 的 -NNX 后缀(如 TS...001-10X) -> (base='TS...001', factor=10.0)。
    无后缀 -> (sid, 1.0)。复用 _DILUTION_RE，与文件名 -NNX 同语义（内容稀释段，供作稀释源）。"""
    m = _DILUTION_RE.search(sid or '')
    if not m:
        return sid, 1.0
    return sid[:m.start()], float(m.group(1))


# 非ICP多样品报告(如PAHS A/B)按 `样品 :` 切段——镜像 ICP 的 `样品识别码：` 切分。
# `样品瓶/样品名称/样品乘积因子` 中间有非空白字符，不匹配(\s* 只吃空白)。
_SAMPLE_MARKER_RE = re.compile(r'样品\s*[:：]\s*')


# ICP-OES 报告（建立者（原始）：ICP）：表头为 分析物/波长/强度/校准浓度(mg/L)/样品浓度(mg/kg)，
# 无「目标化合物」「化合物+保留时间」标记，且一个 PDF 含多个样品(A/B 平行样)。
# 固定表头超集：确保 LIMS 结果列 equipRelativeTitle=校准浓度 能在 _find_result_column 命中。
_ICP_HEADERS = ['分析物', '波长', '强度', '校准浓度', '单位', '样品浓度', '标准偏差', 'RSD']
# 元素符号：Pb/Cd/Cr/Hg/As/Y...（1 大写 + 可选 1 小写）
_ICP_ELEM_RE = re.compile(r'^[A-Z][a-z]?$')


def _detect_icp(content):
    return ('建立者（原始）：ICP' in content) or ('分析物' in content and '校准浓度' in content) \
        or ('分析物' in content and ('平均值数据' in content or '重复测定数据' in content))


def _slice_icp_section(part, marker):
    """从样品段中截取 marker(如「平均值数据：」) 到下一分隔线(----/====)之间的文本。"""
    i = part.find(marker)
    if i < 0:
        return None
    out = []
    for line in part[i + len(marker):].split('\n'):
        if '--------' in line or '========' in line:
            break
        out.append(line)
    return '\n'.join(out)


def _parse_icp_element_line(line):
    """解析 ICP 元素行 -> (分析物, value, status, raw) 或 None。
    行形如 'Pb 220.353 1596.9 0.394 mg/L 19.42 mg/kg'：定位 mg/L，取其前一个 token 为校准浓度(mg/L)。
    部分模板行首带重复序号(如 '1 Pb 220.353 ...')，先剔除纯数字前缀再解析。
    分析物名 = 元素符号+波长(如 'Pb 220.353')，与报告单元格/项目别名一致；负值(低于检出限) -> 未检出。"""
    toks = line.split()
    while toks and toks[0].isdigit():  # 剔除行首重复序号(1/2/3...)
        toks = toks[1:]
    idx = next((i for i, t in enumerate(toks) if t.lower() == 'mg/l'), -1)
    if idx <= 0 or not _ICP_ELEM_RE.match(toks[0]):
        return None
    raw = toks[idx - 1]
    try:
        val = float(raw)
    except ValueError:
        return None
    # 分析物名 = 符号 + 波长(toks[1])，与报告「分析物」列及项目别名匹配
    analyte = toks[0]
    if len(toks) > 1:
        try:
            float(toks[1])  # 确认 toks[1] 是波长而非其它
            analyte = f"{toks[0]} {toks[1]}"
        except ValueError:
            pass
    if val < 0:
        return analyte, 0.0, '未检出', None
    return analyte, val, '检出', raw


def parse_icp_report(content):
    """解析 ICP-OES 报告 -> [(样品标识码, compounds), ...]。
    按「样品识别码：」切段，每段取「平均值数据」(缺则「重复测定数据」)的元素行，校准浓度取 mg/L。
    compounds = {元素: {'value':float,'status':str,'raw':str|None}}，与 parse_pdf_report 同形。"""
    samples = []
    for part in re.split(r'样品识别码：', content)[1:]:  # 第一段为开头 preamble
        m = re.match(r'(\S+)', part)
        sample_id = m.group(1) if m else '?'
        seg = _slice_icp_section(part, '平均值数据：')
        if seg is None:
            seg = _slice_icp_section(part, '重复测定数据：')
        if seg is None:
            continue
        compounds = {}
        for line in seg.split('\n'):
            res = _parse_icp_element_line(line.strip())
            if res and res[0] not in compounds:
                compounds[res[0]] = {'value': res[1], 'status': res[2], 'raw': res[3]}
        if compounds:
            samples.append((sample_id, compounds))
    return samples


# ICP-MS 定量报告(外标法)：英文 Quantitation Report，表头 Element/质量数/.../浓度/单位(ug/L)/...，
# 一个 PDF 一个样品(Sample Name)。固定表头超集：确保 LIMS 结果列 equipRelativeTitle=浓度 能命中。
_ICP_MS_HEADERS = ['Element', '质量数', '内标', '调谐模式', '浓度', '单位', 'RSD(%)', 'CPS', '比率', '检测器模式', 'Time(sec)', 'Rep']


def _detect_icp_ms(content):
    return ('Operator Name ICPMS' in content) or \
        ('FullQuant Table' in content and '质量数' in content and 'ug/L' in content)


def parse_icp_ms_report(content):
    """解析 ICP-MS 定量报告 -> [(样品标识码, compounds), ...]。
    单样品(按 'Sample Name' 取标识码)；元素行 '元素 质量数 调谐 浓度 ug/L ...'，
    浓度取 ug/L 前一个 token；分析物名 = 元素+质量数(如 'Cu 63')。"""
    sid = 'A'
    m = re.search(r'Sample Name\s+(\S+)', content)
    if m:
        sid = m.group(1)
    compounds = {}
    for line in content.split('\n'):
        toks = line.split()
        if len(toks) < 4 or not _ICP_ELEM_RE.match(toks[0]):
            continue
        try:
            int(toks[1])  # 质量数(同位素)
        except ValueError:
            continue
        idx = next((i for i, t in enumerate(toks) if t.lower() == 'ug/l'), -1)
        if idx <= 0:
            continue
        raw = toks[idx - 1]
        try:
            val = float(raw)
        except ValueError:
            continue
        analyte = f"{toks[0]} {toks[1]}"
        if analyte in compounds:
            continue
        if val <= 0:
            compounds[analyte] = {'value': 0.0, 'status': '未检出', 'raw': None}
        else:
            compounds[analyte] = {'value': val, 'status': '检出', 'raw': raw}
    return [(sid, compounds)] if compounds else []


def parse_pdf_report_meta(file_path, field='样品初始质量'):
    """从报告 PDF 提取每样品段的指定数值元数据字段(默认样品初始质量)。
    返回 [(样品标识码, 值字符串|None), ...]，按文档出现顺序(对应平行槽 0/1/...)。
    按「样品识别码：」切段(同 parse_icp_report)，段内正则 `字段[:：]\\s*数值` 取首个数值串；
    无该字段段值为 None。返回报告原始数值串(如 0.3100，保留末尾0)；格式化由调用方按 decimal_places 做。"""
    content = PDFTextExtractor().extract_text_from_pdf(file_path)
    pat = re.compile(r'%s\s*[:：]\s*([0-9]+(?:\.[0-9]+)?)' % re.escape(field))
    out = []
    for part in re.split(r'样品识别码：', content)[1:]:
        m = re.match(r'(\S+)', part)
        sid = m.group(1) if m else '?'
        mv = pat.search(part)
        out.append((sid, mv.group(1) if mv else None))
    return out


_content_id_cache = {}   # ponytail: (path, mtime) -> [base_id]；避免每次关联重抽PDF文本
_icp_ms_name_cache = {}  # ponytail: (path, mtime) -> Sample Name('' 表示非ICP-MS或无)；同上


def extract_icp_ms_sample_name(file_path):
    """ICP-MS 定量报告的 Sample Name(单样品标识码)。非 ICP-MS / 无 Sample Name 返回 None。
    供按报告内容(而非文件名)关联样品——ICP-MS 文件名可能与报告 Sample Name 不一致。
    带 (path, mtime) 缓存，避免每次重抽 PDF 文本。"""
    try:
        mtime = os.path.getmtime(file_path)
    except OSError:
        return None
    key = (file_path, mtime)
    cached = _icp_ms_name_cache.get(key)
    if cached is not None:
        return cached or None
    try:
        content = PDFTextExtractor().extract_text_from_pdf(file_path)
    except Exception:
        _icp_ms_name_cache[key] = ''
        return None
    name = None
    if _detect_icp_ms(content):
        m = re.search(r'Sample Name\s+(\S+)', content)
        if m:
            name = m.group(1)
    _icp_ms_name_cache[key] = name or ''
    return name


def extract_content_sample_ids(file_path):
    """扫描 PDF 内容的 `样品 :` 字段 -> [base_sample_id, ...]（去 -NNX 稀释后缀，保留末尾字母 A/B）。
    供按内容(而非文件名)关联 PDF 与样品编号。失败/无标记返回 []。"""
    try:
        mtime = os.path.getmtime(file_path)
    except OSError:
        return []
    key = (file_path, mtime)
    cached = _content_id_cache.get(key)
    if cached is not None:
        return cached
    try:
        content = PDFTextExtractor().extract_text_from_pdf(file_path)
    except Exception:
        _content_id_cache[key] = []
        return []
    ids = []
    for part in _SAMPLE_MARKER_RE.split(content)[1:]:
        m = re.match(r'(\S+)', part)
        if not m:
            continue
        base, _f = _split_content_dilution(m.group(1))
        if base and base not in ids:
            ids.append(base)
    _content_id_cache[key] = ids
    return ids


def _wrap_compounds(raw):
    """{名: TaggedConcentration} -> {名: {'value','status','raw'}} (parse_pdf_report 同形)。"""
    return {name: {'value': float(c), 'status': getattr(c, 'status', '检出'),
                   'raw': getattr(c, 'raw', None)}
            for name, c in raw.items()}


def _split_non_icp_sections(content, analyzer, file_path):
    """按 `样品 :` 切非ICP报告 -> [(sample_id, compounds), ...]。镜像 parse_icp_report 的 `样品识别码：` 切分。
    无该标记(PAE/氯苯/单样品) -> 回退 [('A', 全content化合物)] 保留旧行为。"""
    parts = _SAMPLE_MARKER_RE.split(content)[1:]
    basename = os.path.basename(file_path)
    if not parts:
        raw, _ = analyzer.parse_report_content(content, basename, file_path)
        return [('A', _wrap_compounds(raw))]
    out = []
    for part in parts:
        m = re.match(r'(\S+)', part)
        sid = m.group(1) if m else '?'
        # ponytail: 复用整个 strategy cascade；每段独立解析，A/B 各自只看本段。
        # TSCA 多样品种PDF(仓库无)此处会看整PDF表格——若出现按页范围限制。
        raw, _ = analyzer.parse_report_content(part, basename, file_path)
        if raw:
            out.append((sid, _wrap_compounds(raw)))
    if out:
        return out
    raw, _ = analyzer.parse_report_content(content, basename, file_path)   # 全段解析失败→回退全content
    return [('A', _wrap_compounds(raw))]


def parse_pdf_report_multi(file_path):
    """无GUI解析谱图PDF -> (samples, headers)。
    samples = [(样品标识码|'A', compounds), ...]：ICP 多样品(A/B)，其它报告单样品 ('A', compounds)。
    compounds = {名: {'value':float,'status':str,'raw':str|None}}，status ∈ {'检出','未检出','未校正'}；
    headers = 报告表头 token 列表（ICP 固定含「校准浓度」等），供按 equipRelativeTitle 匹配选列。"""
    analyzer = _HeadlessReportAnalyzer()
    content = analyzer.pdf_extractor.extract_text_from_pdf(file_path)
    if _detect_icp_ms(content):
        return parse_icp_ms_report(content), list(_ICP_MS_HEADERS)
    if _detect_icp(content):
        return parse_icp_report(content), list(_ICP_HEADERS)
    return _split_non_icp_sections(content, analyzer, file_path), _extract_column_headers(content)


def filter_samples_by_code(samples, sample_code):
    """从 parse_pdf_report_multi 的 samples=[(样品识别码, compounds|值), ...] 中，只保留属于
    sample_code 的段。ICP 一份 PDF 可能含多个报验批样品——样品识别码形如 TN26070722001A
    (报验编号+3位小号+平行后缀A/B)，去末尾字母得 base=TN26070722001。

    匹配用前缀而非精确相等：sample_code 可能只到报验编号(无小号，如目录+单PDF 时由文件名
    TN26070722.pdf 解析得 sc='TN26070722')，也可能含小号('TN26070722001')。前缀匹配两种都成立，
    且仍能把 21 批(TN26070721001A)与 22 批(TN26070722001A)区分开(报验编号不同)。

    base 为空(如非 ICP 报告的 'A'、'Blank'/'STD' 等纯字母识别码)不参与匹配，避免 ''.startswith 误命中。
    无任何段匹配时原样返回：兼容非 ICP 报告(parse_pdf_report_multi 对单样品报告返回 [('A', ...)])，
    避免回归。对 parse_pdf_report_meta 的 [(sid, 值), ...] 同样适用。"""
    def _base(sid):
        s = _DILUTION_RE.sub('', (sid or '').strip())   # 去 -NNX(内容稀释段 TS...001-10X)
        return re.sub(r'[A-Za-z]+$', '', s)
    target = (sample_code or '').strip()
    if not target:
        return samples
    matched = []
    for s in samples:
        sb = _base(s[0])
        if sb and (sb.startswith(target) or target.startswith(sb)):
            matched.append(s)
    return matched if matched else samples


def parse_pdf_report(file_path):
    """无GUI解析单个谱图PDF -> (compounds, headers)。多样品报告取首个样品（向后兼容旧调用/自检/稀释路径）。"""
    samples, headers = parse_pdf_report_multi(file_path)
    return (samples[0][1] if samples else {}), headers


# 运行应用程序
if __name__ == "__main__":
    import sys
    # ponytail: 自检——稀释倍数提取
    assert _dilution_factor('TS24031359001A-50X.pdf') == 50.0
    assert _dilution_factor('TS24031359001A-2x.pdf') == 2.0
    assert _dilution_factor('TS24031359001A.pdf') == 1.0
    # ponytail: 自检——多报验批样品段按 sampleCode 过滤(21/22 批共一份 PDF)
    _src = [('TN26070721001A', {}), ('TN26070721001B', {}),
            ('TN26070722001A', {}), ('TN26070722001B', {})]
    _fs = filter_samples_by_code(_src, 'TN26070722001')  # sc 含小号
    assert len(_fs) == 2 and _fs[0][0] == 'TN26070722001A' and _fs[1][0] == 'TN26070722001B', _fs
    _fs2 = filter_samples_by_code(_src, 'TN26070722')  # sc 只到报验编号(文件名解析)
    assert len(_fs2) == 2 and _fs2[0][0] == 'TN26070722001A', _fs2
    assert filter_samples_by_code([('A', {})], 'TN26070722001') == [('A', {})]  # 非 ICP 回退
    assert filter_samples_by_code([('Blank', {})], 'TN26070722001') == [('Blank', {})]  # 空基不误匹配
    # ponytail: 自检——内容稀释段 -NNX 不再被 _base 误剥成 ...-10；A/B 末字母仍剥
    assert filter_samples_by_code([('TS26072901001-10X', {})], 'TS26072901001') == [('TS26072901001-10X', {})]
    assert _split_content_dilution('TS26072901001-10X') == ('TS26072901001', 10.0)
    assert _split_content_dilution('TS26080100001A') == ('TS26080100001A', 1.0)
    # ponytail: 自检——ICP 样品初始质量提取(三份报告存在时；不存在则跳过)
    import glob as _glob, os as _os

    def _is_num(s):
        try:
            float(s); return True
        except (TypeError, ValueError):
            return False
    for _icp in sorted(_glob.glob('谱图/报告解析/ICP/TN*.pdf')):
        _meta = parse_pdf_report_meta(_icp)
        _vals = [v for _s, v in _meta if v]
        assert len(_meta) >= 2 and all(map(_is_num, _vals)), (_icp, _meta)
        print(f"[meta 样品初始质量] {_os.path.basename(_icp)}: {_meta}")
    # ponytail: 自检——GC-FID(cid 乱码)动态解码：化合物名无 cid 残留(字体解码成功) + status 合法。
    # 不写死面板名单——只校验「解码出可读中文 + 产出化合物」，面板随报告动态变化。
    for _gc in sorted(set(_glob.glob('谱图/GC/*.pdf') + _glob.glob('谱图/GC/*.PDF'))):
        _s, _ = parse_pdf_report_multi(_gc)
        assert _s, f"GC-FID 未解析到化合物: {_gc}"
        for _sid, _c in _s:
            assert _c, f"GC-FID 样品无化合物: {_gc}"
            for _n, _v in _c.items():
                assert '(cid:' not in _n, (_gc, _n)          # 系统字体解码成功，无 cid 残留
                assert _v['status'] in ('检出', '未检出'), (_gc, _n, _v)
        print(f"[GC-FID] {_os.path.basename(_gc)}: {sum(len(c) for _, c in _s)} 个化合物")
    # ponytail: 自检——ICP-MS 定量报告(ug/L)：含 'Cu 63' 等元素+质量数，浓度>0，status 合法
    for _ms in sorted(_glob.glob('谱图/报告解析/ICP/TS*-定量报告-内标.pdf')):
        _s, _h = parse_pdf_report_multi(_ms)
        assert _s and _s[0][1], f"ICP-MS 未解析到化合物: {_ms}"
        for _n, _v in _s[0][1].items():
            assert ' ' in _n and _v['status'] in ('检出', '未检出'), (_ms, _n, _v)
        assert '浓度' in _h, (_ms, _h)
        print(f"[ICP-MS] {_os.path.basename(_ms)}: {len(_s[0][1])} 个化合物，样品 {_s[0][0]}")
    # ponytail: 自检——带参数则解析该PDF并打印各样品化合物+status(断言至少1个)；无参数启动GUI。
    if len(sys.argv) > 1:
        _samples, _h = parse_pdf_report_multi(sys.argv[1])
        assert _samples, f"未解析到化合物: {sys.argv[1]}"
        for _sid, _c in _samples:
            print(f"[样品 {_sid}] {len(_c)} 个化合物；表头含校准浓度={'校准浓度' in _h}")
            for _n, _v in list(_c.items()):
                print(f"  {_n}\t{_v['value']}\t{_v['status']}\t{_v['raw']}")
        print(f"共 {len(_samples)} 个样品；表头={_h}")
    else:
        # 配置ttk样式为现代主题
        try:
            from ctypes import windll

            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass

        app = ChemicalReportAnalyzer()
        app.run()
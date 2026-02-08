# Droplet Impact Toolkit

> Data processing toolkit for droplet impact dynamics research - supports image analysis, force calibration, and non-dimensional scaling (We, F*, t/τ).

本工具集用于液滴撞击实验的数据采集、处理和分析，包含三个独立的Python工具。

---

## 环境要求

### Python版本
- Python 3.8+

### 依赖安装

```bash
# 基础依赖
pip install numpy pandas openpyxl

# 图像测量工具依赖
pip install PySide6 Pillow

# 绘图工具依赖
pip install plotly dash

# 峰值分析工具依赖
pip install PySide6 matplotlib scipy
```

或使用conda一键安装：
```bash
conda install -c conda-forge pyside6 pillow openpyxl numpy pandas plotly dash matplotlib scipy -y
```

---

## 工具一：液滴直径测量工具

**文件**: `droplet_diameter_tool.py`

### 功能概述
从高速摄像机拍摄的图像序列中手动测量液滴的水平直径(Dh)和垂直直径(Dv)，并计算等效直径D0。

### 启动方式
```bash
python droplet_diameter_tool.py
```

### 文件夹结构要求
```
根目录/
├── 浓度1_速度1/
│   ├── image001.png
│   ├── image002.png
│   └── ...
├── 浓度2_速度2/
│   ├── image001.png
│   └── ...
└── ...
```
- 子文件夹命名格式：`浓度_速度_其他信息`（下划线分隔）
- 支持的图像格式：PNG, JPG, JPEG, BMP

### 操作流程

1. **加载数据**
   - 输入或浏览选择根目录
   - 点击"Load"加载所有条件文件夹

2. **校准（每个条件需单独校准）**
   - 输入针头外径（默认0.41mm）
   - 点击"1) Calibrate Needle"
   - 在图像上点击针头两端（水平方向锁定）
   - 完成后显示 mm/px 换算系数

3. **测量直径**
   - 点击"2) Measure Dh"测量水平直径
   - 点击"3) Measure Dv"测量垂直直径
   - 两者完成后自动计算 D0 = (Dh² × Dv)^(1/3)

4. **时间轴设置**
   - 设置FPS（帧率）
   - 导航到撞击起始帧，点击"Set current image as T0"
   - 之后所有帧会显示相对时间

5. **导出结果**
   - 点击"Export Excel"导出所有测量结果
   - 包含Summary汇总表和各条件详细数据表

### 快捷操作
| 操作 | 说明 |
|------|------|
| Ctrl + 滚轮 | 缩放图像 |
| Ctrl + 左键拖动 | 平移图像 |
| Ctrl + Z | 撤销当前测量点 |
| Esc | 重置当前测量步骤 |
| Ctrl + A | 全选表格行 |
| Delete | 删除选中的表格行 |

### 输出字段说明
| 字段 | 说明 |
|------|------|
| Dh(mm) | 水平直径 |
| Dv(mm) | 垂直直径 |
| D0(mm) | 等效直径 = (Dh² × Dv)^(1/3) |
| Time_s | 相对T0的时间（秒） |
| Frame | 帧序号 |

---

## 工具二：数据绘图与无量纲校正工具

**文件**: `figure_tool.py`

### 功能概述
读取Excel数据文件，绘制双对数图，支持无量纲参数校正和Origin格式导出。

### 启动方式
```bash
python figure_tool.py
# 或指定文件
python figure_tool.py --excel "数据文件.xlsx" --port 8050
```

启动后在浏览器访问 `http://127.0.0.1:8050`

### 主要功能

1. **多工作表绘图**
   - 支持同时选择多个工作表
   - 自动识别数值列
   - 双对数坐标轴

2. **绘图模式**
   - **Raw（原始）**: 显示所有数据点
   - **Mean + Error bars**: 按X值分组计算均值和误差棒
   - 误差类型：STD（标准差）、SEM（标准误）、95% CI（置信区间）

3. **无量纲参数校正**
   - 支持6参数校正模型
   - 参数：rho_old, sigma_old, d_old_mm, rho_new, sigma_new, d_new_mm
   - 自动重算 We, F*, t/τρ, t/τγ 等无量纲量

4. **数据导出**
   - **Export plotted data**: 导出当前图表数据（Origin格式 + 统计表）
   - **Export corrected workbook**: 导出校正后的完整工作簿（保留原格式）

### 预期Excel列名
| 列名 | 说明 |
|------|------|
| We | 韦伯数 |
| U0 (m/s) | 撞击速度 |
| D0 (mm) | 液滴直径 |
| F1*, F2* | 无量纲力峰值 |
| t1/τρ, t2/τρ | 无量纲时间（惯性时间尺度） |
| t1/τγ, t2/τγ | 无量纲时间（毛细时间尺度） |

---

## 工具三：峰值分析工具

**文件**: `peak_analysis_tool.py`

### 功能概述
分析力传感器采集的CSV数据，进行滤波、校准、峰值提取，计算无量纲参数。

### 启动方式
```bash
python peak_analysis_tool.py
```

### CSV文件格式要求
- 前9行为头信息（自动跳过）
- 第1列：时间（秒）
- 第2列：电压（V）
- 文件命名格式：`流体名称_速度_直径_日期_时间.csv`
  - 例：`2000ppm_2.45_0.422_20251223_154729.csv`

### 操作流程

1. **加载数据**
   - 浏览或输入CSV文件路径
   - 点击"加载数据"
   - 程序自动从文件名提取流体类型、U0、D0

2. **滤波设置**
   - **FIR低通滤波**（默认）：设置截止频率（默认3000Hz）
   - **高斯滤波**（可选）：设置Sigma值
   - 点击"应用滤波"更新

3. **T0校准**
   - 方法一：在"FIR低通滤波曲线"图中点击快速定位峰值
   - 方法二：点击"精确校准T0"，在"校准曲线（滤波）"图中精确选点
   - 方法三：使用"左移/右移"按钮微调（可设置步长）

4. **峰值提取**
   - **自动提取**：点击"自动提取"自动识别Peak1和Peak2
   - **手动选择**：点击"点选Peak1/Peak2"在图中手动选择
   - 支持微调按钮精细调整峰值位置

5. **查看结果**
   - 结果区显示：t/τρ, t/τγ, F*, We
   - 无量纲曲线图同时显示两种时间尺度

6. **导出数据**
   - **导出数据**：导出校准曲线和无量纲曲线到Excel
   - **导出峰值**：追加峰值数据到汇总文件

### 内置流体参数
| 流体 | 密度 ρ (kg/m³) | 表面张力 σ (N/m) |
|------|----------------|------------------|
| water | 997.0 | 0.0707 |
| 800ppm | 998.0 | 0.0707 |
| 2000ppm | 998.0 | 0.072 |
| 4000ppm | 998.0 | 0.072 |
| 6000ppm | 998.0 | 0.072 |
| 10000ppm | 1002.0 | 0.074 |

### 物理公式
| 参数 | 公式 |
|------|------|
| 韦伯数 We | ρU₀²D₀ / σ |
| 无量纲力 F* | F / (ρU₀²D₀²) |
| 惯性时间尺度 τρ | D₀ / U₀ |
| 毛细时间尺度 τγ | √(ρD₀³/σ) |
| 碰撞力 F | 0.05 × (V - V₀) |

### 输出文件结构
导出的Excel包含以下工作表：
1. **滤波校准曲线**：相对时间 vs 碰撞力（滤波后）
2. **无量纲曲线_τρ**：t/τρ vs F*
3. **无量纲曲线_τγ**：t/τγ vs F*
4. **原始校准曲线**：相对时间 vs 碰撞力（原始）

---

## 典型工作流程

```
1. 高速摄像 → droplet_diameter_tool.py → 测量D0
                                            ↓
2. 力传感器采集 → peak_analysis_tool.py → 提取峰值、计算无量纲参数
                                            ↓
3. 汇总数据 → figure_tool.py → 绘图、校正、导出
```

---

## 常见问题

### Q: 图像加载失败
A: 检查图像格式是否为PNG/JPG/BMP，路径是否包含特殊字符。

### Q: CSV读取报错
A: 确认CSV文件前9行为头信息，数据从第10行开始，使用逗号分隔。

### Q: 绘图工具无法启动
A: 检查端口8050是否被占用，可使用 `--port 8051` 指定其他端口。

### Q: 无量纲校正不生效
A: 确保6个参数（rho_old, sigma_old, d_old_mm, rho_new, sigma_new, d_new_mm）全部填写。

---

## 版本信息

- droplet_diameter_tool.py: Manual + Timeline + Overlay
- figure_tool.py: Plotly + Dash + Non-dimensional Correction
- peak_analysis_tool.py: v3.4 (自动加载优化版)

### 更新日志

#### v3.4 (2026-02-09)
- **优化**: 数据加载流程改进
  - 选择流体类型后自动加载第一个速度的数据
  - 无需手动点击速度下拉框即可查看数据
  - 提升单速度场景下的用户体验

#### v3.3 (2026-02-06)
- **修复**: 批量导出对比数据时X轴范围截断问题
  - 原逻辑取所有曲线时间范围的**交集**，导致边界较长的曲线数据被截断
  - 现改为取**并集**，完整保留所有曲线的原始范围
  - 超出某条曲线范围的Y值自动填充为空值
  - 优化插值点数计算，基于最密集曲线的密度确保分辨率

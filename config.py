import os

# ================= 1. 品牌信息配置 =================
APP_NAME = "极简票流 (FlowInvoice)"
APP_SUBTITLE = "智能发票提取与排版引擎"
PAGE_TITLE = "极简票流 FlowInvoice"

# ================= 2. 目录与路径配置 =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER_NAME = "初始发票箱"  
INPUT_FOLDER = os.path.join(BASE_DIR, INPUT_FOLDER_NAME)

# ================= 3. 财务与算税规则 (2024法定抵扣标准) =================
HSR_TAX_RATE = 1.09    # 高铁/火车票 (9%)
FLIGHT_TAX_RATE = 1.09 # 飞机行程单 (9% 仅票价+燃油)
TAXI_TAX_RATE = 1.03   # 出租车/公交定额票 (3%)

# ================= 4. 排版与识别参数 =================
MERGE_SCALE_FACTOR = 0.94  # 打印安全边距缩放，防止打印机物理裁切
MAX_TAXI_AMOUNT = 2000      # 打车票金额上限，超出视为 OCR 误读（针式打印极度模糊时会把税号识别成金额）

# ================= 5. 识别置信度阈值 =================
CONFIDENCE_HIGH_THRESHOLD = 80   # >= 80 不高亮（识别可信）
CONFIDENCE_LOW_THRESHOLD  = 50   # <  50 红色高亮（需人工核查）；50~79 黄色提示（建议核对）
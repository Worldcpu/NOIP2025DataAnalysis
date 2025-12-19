import pandas as pd
import numpy as np

# =============================
# 1. 读取 Excel 数据
# =============================
file_path = "yd.xlsx"  # 修改为你的文件路径
sheet_name = 0
df = pd.read_excel(file_path, sheet_name=sheet_name)

# =============================
# 2. 提取题目列
# =============================
question_cols = [col for col in df.columns if col.startswith('#')]
X = df[question_cols].copy()

# =============================
# 3. 将非数值转换为数值，无法转换的设为0
# =============================
for col in X.columns:
    X[col] = pd.to_numeric(X[col], errors='coerce').fillna(0)

X_values = X.values.astype(float)

# 样本量和题目数
n_items = X_values.shape[1]
n_subjects = X_values.shape[0]

# =============================
# 4. 计算每题方差和总分方差
# =============================
item_variances = X_values.var(axis=0, ddof=1)
total_scores = X_values.sum(axis=1)
total_variance = total_scores.var(ddof=1)

# =============================
# 5. Cronbach's alpha
# =============================
cronbach_alpha = n_items / (n_items - 1) * (1 - item_variances.sum() / total_variance)

# =============================
# 6. 标准测量误差 SEM
# =============================
sd_total = np.sqrt(total_variance)
sem = sd_total * np.sqrt(1 - cronbach_alpha)

# =============================
# 7. 输出结果
# =============================
print(f"样本量: {n_subjects}")
print(f"题目数: {n_items}")
print(f"Cronbach's α: {cronbach_alpha:.4f}")
print(f"总分 SD: {sd_total:.4f}")
print(f"标准测量误差 SEM: {sem:.4f}")

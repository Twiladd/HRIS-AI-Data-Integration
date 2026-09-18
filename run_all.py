import subprocess
import sys
from pathlib import Path


# =====================================
# 1. 获取项目根目录
# =====================================

BASE_DIR = Path(__file__).resolve().parent


# =====================================
# 2. 定义需要执行的程序
# =====================================

steps = [
    {
        "name": "AI 字段映射分析",
        "file": BASE_DIR / "src" / "ai_mapper.py"
    },
    {
        "name": "员工数据清洗与校验",
        "file": BASE_DIR / "src" / "main.py"
    },
    {
        "name": "AI 错误分析",
        "file": BASE_DIR / "src" / "ai_error_analyzer.py"
    }
]


# =====================================
# 3. 运行单个 Python 程序
# =====================================

def run_step(name, file_path):

    print()
    print("=" * 60)
    print(f"开始：{name}")
    print("=" * 60)

    if not file_path.exists():

        print(f"错误：找不到文件：{file_path}")

        return False

    result = subprocess.run(
        [sys.executable, str(file_path)],
        cwd=BASE_DIR
    )

    if result.returncode != 0:

        print()
        print("=" * 60)
        print(f"失败：{name}")
        print("=" * 60)

        return False

    print()
    print("=" * 60)
    print(f"完成：{name}")
    print("=" * 60)

    return True


# =====================================
# 4. 依次运行
# =====================================

print()
print("=" * 60)
print("HRIS AI 自动化工作流开始")
print("=" * 60)


for step in steps:

    success = run_step(
        step["name"],
        step["file"]
    )

    if not success:

        print()
        print("工作流已停止，请检查上面的错误信息。")

        sys.exit(1)


# =====================================
# 5. 检查最终输出
# =====================================

output_dir = BASE_DIR / "output"

expected_files = [
    "ai_mapping_suggestion.xlsx",
    "hris_import.xlsx",
    "sync_errors.xlsx",
    "ai_error_analysis.xlsx"
]


print()
print("=" * 60)
print("最终输出检查")
print("=" * 60)


for filename in expected_files:

    file_path = output_dir / filename

    if file_path.exists():

        print(f"[OK] {filename}")

    else:

        print(f"[MISSING] 缺少：{filename}")


print()
print("=" * 60)
print("HRIS AI 自动化工作流完成")
print("=" * 60)
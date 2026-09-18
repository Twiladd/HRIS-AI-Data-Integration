from pathlib import Path

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel


# =========================================
# 1. 创建 FastAPI 应用
# =========================================

app = FastAPI(
    title="Mock HRIS API",
    description="用于学习 HRIS 数据同步的本地模拟系统",
    version="1.0.0"
)


# =========================================
# 2. 模拟数据库
# =========================================

DATA_DIR = Path(__file__).resolve().parent / "data"

DATA_DIR.mkdir(exist_ok=True)

DATABASE_FILE = DATA_DIR / "mock_hris_database.json"


def load_database():

    if not DATABASE_FILE.exists():
        return {}

    import json

    with open(
        DATABASE_FILE,
        "r",
        encoding="utf-8"
    ) as file:

        return json.load(file)


def save_database(database):

    import json

    with open(
        DATABASE_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            database,
            file,
            ensure_ascii=False,
            indent=2
        )


# =========================================
# 3. 员工数据结构
# =========================================

class Employee(BaseModel):

    employee_id: str
    full_name: str
    gender_code: str
    department_code: str
    hire_date: str
    work_email: str | None = None
    employment_status: str


# =========================================
# 4. 首页
# =========================================

@app.get("/")
def root():

    return {
        "system": "Mock HRIS",
        "status": "running"
    }


# =========================================
# 5. 获取全部员工
# =========================================

@app.get("/employees")
def get_employees():

    database = load_database()

    return {
        "count": len(database),
        "employees": list(database.values())
    }


# =========================================
# 6. 获取单个员工
# =========================================

@app.get("/employees/{employee_id}")
def get_employee(employee_id: str):

    database = load_database()

    if employee_id not in database:

        raise HTTPException(
            status_code=404,
            detail="Employee not found"
        )

    return database[employee_id]


# =========================================
# 7. 新增员工
# =========================================

@app.post("/employees")
def create_employee(employee: Employee):

    database = load_database()

    employee_id = employee.employee_id

    # 检查重复
    if employee_id in database:

        raise HTTPException(
            status_code=409,
            detail=f"Employee {employee_id} already exists"
        )

    # 保存
    database[employee_id] = employee.model_dump()

    save_database(database)

    return {
        "success": True,
        "message": "Employee created",
        "employee_id": employee_id
    }
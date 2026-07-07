#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据 PostgreSQL 查询结果导出 Excel 文件。

使用方法：
    python generate-data.py

依赖：
    pip install psycopg2-binary openpyxl
    # 或使用 psycopg(v3)：pip install psycopg[binary] openpyxl

说明：
- 数据库连接信息从当前脚本同目录的 db.env 读取，支持 `key: value` 或 `key=value`。
- SQL 文件按 utf-8/gb18030 等编码自动尝试读取。
- Excel 第一行列标题使用 SQL 返回的字段名。
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple


BASE_DIR = Path(__file__).resolve().parent
LOCAL_LIB_DIR = BASE_DIR / ".pythonlibs"
if LOCAL_LIB_DIR.exists():
    sys.path.insert(0, str(LOCAL_LIB_DIR))

try:
    from openpyxl import Workbook
except ImportError:  # 运行时给出更清晰的依赖提示
    Workbook = None  # type: ignore


DB_ENV_FILE = BASE_DIR / "db.env"

EXPORT_TASKS: Sequence[Tuple[str, str]] = (
    ("main_inventor.sql", "库存数据.xlsx"),
    ("dfn_sales_prd_summary.sql", "销售数据.xlsx"),
)


class ConfigError(RuntimeError):
    """数据库配置错误。"""


class DependencyError(RuntimeError):
    """缺少 Python 依赖。"""


def read_text_auto(path: Path) -> str:
    """读取文本文件，兼容 UTF-8 和常见中文 Windows 编码。"""
    data = path.read_bytes()
    encodings = ("utf-8-sig", "utf-8", "gb18030", "gbk", "cp936")

    last_error: Exception | None = None
    for encoding in encodings:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError as exc:
            last_error = exc

    raise UnicodeDecodeError(
        "unknown",
        data,
        0,
        1,
        f"无法识别文件编码：{path}；最后错误：{last_error}",
    )


def load_db_config(path: Path = DB_ENV_FILE) -> Dict[str, Any]:
    """从 db.env 加载 PostgreSQL 连接配置。"""
    if not path.exists():
        raise ConfigError(f"找不到数据库配置文件：{path}")

    config: Dict[str, Any] = {}
    for line_no, raw_line in enumerate(read_text_auto(path).splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        if "=" in line:
            key, value = line.split("=", 1)
        elif ":" in line:
            key, value = line.split(":", 1)
        else:
            raise ConfigError(f"{path} 第 {line_no} 行格式错误，应为 key=value 或 key: value")

        key = key.strip().lower()
        value = value.strip().strip('"').strip("'")
        config[key] = value

    # 兼容常见字段名
    if "username" in config and "user" not in config:
        config["user"] = config["username"]
    if "dbname" in config and "database" not in config:
        config["database"] = config["dbname"]

    required = ("host", "port", "database", "user", "password")
    missing = [key for key in required if not config.get(key)]
    if missing:
        raise ConfigError(f"{path} 缺少必要配置项：{', '.join(missing)}")

    try:
        config["port"] = int(config["port"])
    except ValueError as exc:
        raise ConfigError(f"数据库端口不是有效数字：{config['port']!r}") from exc

    return {
        "host": config["host"],
        "port": config["port"],
        "database": config["database"],
        "user": config["user"],
        "password": config["password"],
    }


def connect_postgres(config: Dict[str, Any]):
    """连接 PostgreSQL，优先使用 psycopg2，未安装时尝试 psycopg(v3)。"""
    try:
        import psycopg2  # type: ignore

        return psycopg2.connect(**config)
    except ImportError:
        try:
            import psycopg  # type: ignore

            # psycopg(v3) 使用 dbname 参数；同时兼容 database 写法。
            config_v3 = dict(config)
            config_v3["dbname"] = config_v3.pop("database")
            return psycopg.connect(**config_v3)
        except ImportError as exc:
            raise DependencyError(
                "缺少 PostgreSQL 驱动，请安装：pip install psycopg2-binary\n"
                "或安装 psycopg(v3)：pip install 'psycopg[binary]'"
            ) from exc


def get_column_names(description: Sequence[Any] | None) -> List[str]:
    """从 cursor.description 中提取列名。"""
    if not description:
        raise RuntimeError("SQL 没有返回结果列，无法生成 Excel。")

    names: List[str] = []
    for item in description:
        # psycopg2: tuple，第 0 项为列名；psycopg3: Column 对象有 name 属性。
        name = getattr(item, "name", None)
        if name is None:
            name = item[0]
        names.append(str(name))
    return names


def normalize_cell_value(value: Any) -> Any:
    """把少数 openpyxl 不支持的值转换成可写入 Excel 的值。"""
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (bytes, bytearray, memoryview)):
        return bytes(value).hex()
    # openpyxl 支持 datetime/date/time/Decimal 等常见类型；其它类型转字符串更稳妥。
    module = type(value).__module__
    if module.startswith("datetime") or module.startswith("decimal"):
        return value
    return str(value)


def export_query_to_excel(conn: Any, sql_path: Path, output_path: Path, fetch_size: int = 5000) -> int:
    """执行 SQL 并将结果导出为 Excel，返回导出的数据行数。"""
    if not sql_path.exists():
        raise FileNotFoundError(f"找不到 SQL 文件：{sql_path}")

    sql = read_text_auto(sql_path).strip()
    if not sql:
        raise RuntimeError(f"SQL 文件为空：{sql_path}")

    if Workbook is None:
        raise DependencyError("缺少 Excel 依赖，请安装：pip install openpyxl")

    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet(title="Sheet1")

    row_count = 0
    with conn.cursor() as cursor:
        cursor.execute(sql)
        worksheet.append(get_column_names(cursor.description))

        while True:
            rows = cursor.fetchmany(fetch_size)
            if not rows:
                break
            for row in rows:
                worksheet.append([normalize_cell_value(value) for value in row])
                row_count += 1

    workbook.save(output_path)
    return row_count


def main() -> int:
    try:
        config = load_db_config()
        print(f"正在连接 PostgreSQL：{config['host']}:{config['port']}/{config['database']}")

        conn = connect_postgres(config)
        try:
            for sql_file, output_file in EXPORT_TASKS:
                sql_path = BASE_DIR / sql_file
                output_path = BASE_DIR / output_file
                print(f"正在执行 {sql_file}，导出到 {output_file} ...")
                row_count = export_query_to_excel(conn, sql_path, output_path)
                print(f"完成：{output_file}，共导出 {row_count} 行数据。")
        finally:
            close = getattr(conn, "close", None)
            if close is not None:
                close()

        print("全部导出完成。")
        return 0
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

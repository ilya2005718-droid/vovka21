import pymysql
import psycopg2
import pyodbc
from typing import Dict, Any
from core import config

def get_mysql_conn_params() -> Dict[str, Any]:
    db = config.get_db_settings()
    mysql = db.get("mysql", {})
    return {
        "host": mysql.get("host", "192.168.144.221"),
        "port": int(mysql.get("port", 3306)),
        "user": mysql.get("user", "ascue"),
        "password": mysql.get("password", "ascue321"),
        "db": mysql.get("database", "ascue"),
        "charset": mysql.get("charset", "utf8"),
        "cursorclass": pymysql.cursors.DictCursor,
    }

def get_mysql_connection():
    params = get_mysql_conn_params()
    return pymysql.connect(**params)

def get_postgres_connection(name: str = "monitor"):
    db = config.get_db_settings()
    connstr = db.get(name)
    if not connstr:
        raise RuntimeError(f"No postgres connection string for {name}")
    try:
        return psycopg2.connect(connstr)
    except Exception:
        parts = {}
        for part in connstr.split(";"):
            if "=" in part:
                k, v = part.split("=", 1)
                parts[k.strip().lower()] = v.strip()
        host = parts.get("server") or parts.get("host")
        port = parts.get("port")
        user = parts.get("user id") or parts.get("userid") or parts.get("user")
        password = parts.get("password")
        dbname = parts.get("database") or parts.get("initial catalog")
        conn_info = {
            "host": host,
            "port": int(port) if port else 5432,
            "user": user,
            "password": password,
            "dbname": dbname,
        }
        return psycopg2.connect(**{k: v for k, v in conn_info.items() if v})

def get_mssql_connection():
    db = config.get_db_settings()
    connstr = db.get("put_sphera")
    if not connstr:
        raise RuntimeError("No mssql connection configured (put_sphera)")
    try:
        return pyodbc.connect(connstr)
    except Exception as e:
        raise

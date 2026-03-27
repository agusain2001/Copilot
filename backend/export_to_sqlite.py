import asyncio
import os
import sys
import uuid
import datetime
import json

# Ensure app is importable
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from sqlalchemy.ext.asyncio import create_async_engine
from sqlalchemy import create_engine
from app.database import Base
from app.config import settings

# Import all models
from app.models.user import User
from app.models.user_preference import UserPreference
from app.models.shop import Shop
from app.models.report_type import ReportType
from app.models.report import Report
from app.models.report_sequence import ReportSequence
from app.models.exchange_rate import ExchangeRate

def serialize(val):
    if isinstance(val, uuid.UUID):
        return str(val)
    if isinstance(val, datetime.datetime):
        return val.replace(tzinfo=None) # SQLite expects naive datetimes
    if isinstance(val, (dict, list)):
        return json.dumps(val)
    return val

async def main():
    pg_url = settings.DATABASE_URL
    # We will output the file to the root directory for easy sharing
    sqlite_file_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "shared.db")
    if os.path.exists(sqlite_file_path):
        os.remove(sqlite_file_path)
    sqlite_url = f"sqlite:///{sqlite_file_path}"
    
    print("Initiating PostgreSQL connection...")
    pg_engine = create_async_engine(pg_url)
    
    print(f"Creating SQLite database at {sqlite_file_path}...")
    sqlite_engine = create_engine(sqlite_url)
    
    print("Creating schema in SQLite...")
    Base.metadata.create_all(sqlite_engine)
    
    print("Copying data...")
    import sqlite3
    
    # Connect directly with sqlite3
    sqlite_conn = sqlite3.connect(sqlite_file_path)
    sqlite_cursor = sqlite_conn.cursor()
    
    async with pg_engine.connect() as pg_conn:
        for table in Base.metadata.sorted_tables:
            print(f" -> Exporting table: {table.name}")
            res = await pg_conn.execute(table.select())
            rows = res.fetchall()
            if rows:
                keys = list(res.keys())
                
                # Create raw INSERT SQL
                placeholders = ", ".join(["?"] * len(keys))
                cols = ", ".join(keys)
                insert_sql = f"INSERT INTO {table.name} ({cols}) VALUES ({placeholders})"
                
                insert_data = []
                for row in rows:
                    row_list = [serialize(v) for v in row]
                    insert_data.append(row_list)
                
                sqlite_cursor.executemany(insert_sql, insert_data)
                sqlite_conn.commit()
    
    sqlite_conn.close()
    await pg_engine.dispose()
    print("Export complete! You can now share the 'shared.db' file.")

if __name__ == "__main__":
    asyncio.run(main())

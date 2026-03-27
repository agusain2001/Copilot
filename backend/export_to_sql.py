import sqlite3
import os

db_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "shared.db")
sql_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "shared.sql")

def main():
    if not os.path.exists(db_path):
        print(f"Error: Database {db_path} not found.")
        return
        
    print(f"Dumping {db_path} to {sql_path}...")
    with sqlite3.connect(db_path) as con:
        with open(sql_path, "w", encoding="utf-8") as f:
            for line in con.iterdump():
                f.write(f"{line}\n")
    print("Export complete! SQL file is ready.")

if __name__ == "__main__":
    main()

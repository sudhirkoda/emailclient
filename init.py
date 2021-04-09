from db import Database

db = Database()

createtablesql = """
CREATE TABLE IF NOT EXISTS tbl_emailclient(
id INTEGER PRIMARY KEY AUTOINCREMENT,
"from" TEXT NOT NULL,
"to" TEXT,
cc TEXT,
bcc TEXT,
subject TEXT,
status TEXT NOT NULL,
response TEXT NOT NULL,
emailsenttime DATETIME NOT NULL,
"type" TEXT NOT NULL)
"""

status, message = db.execute(sql=createtablesql)
if status:
    print("DB init successful")
    exit()
print(f"DB init failed, response- {message}")

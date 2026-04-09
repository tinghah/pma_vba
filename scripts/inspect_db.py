import psycopg2
import sys

conn_str = "host=172.23.86.119 port=5432 dbname=pma_hr user=postgres password=Abc123"

try:
    conn = psycopg2.connect(conn_str)
    cur = conn.cursor()
    
    print("Listing tables:")
    cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public';")
    tables = cur.fetchall()
    for t in tables:
        tname = t[0]
        print(f"\nTable: {tname}")
        cur.execute(f"SELECT column_name, data_type FROM information_schema.columns WHERE table_name = '{tname}';")
        cols = cur.fetchall()
        for c in cols:
            print(f"  - {c[0]} ({c[1]})")
            
    cur.close()
    conn.close()
except Exception as e:
    print(f"Error: {e}")

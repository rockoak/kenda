import psycopg2
import sys

def import_perdep_data(db_config, file_path):
    """
    從 perdep.txt 檔案讀取資料，並以 DELETE 後 INSERT 的方式匯入 PostgreSQL。
    """
    try:
        # 連接到 PostgreSQL 資料庫
        db_config['database'] = db_config.pop('dbname')
        conn = psycopg2.connect(**db_config)
        cursor = conn.cursor()

        # 讀取 perdep.txt 檔案，並將編碼改為 'Big5'
        with open(file_path, 'r', encoding='Big5') as f:
            lines = f.readlines()
        
        insert_count = 0
        delete_count = 0
        
        print("開始處理資料並匯入...")
        
        for line in lines:
            line = line.strip()

            if not line:
                continue

            if line.endswith('|'):
                line = line[:-1]

            values = line.split('|')

            if len(values) != 11:
                print(f"警告: 發現無效行，欄位數不符 (應為11，但為 {len(values)})，跳過 -> {line}")
                continue

            subno = values[0]
            serno = values[1]
            depno = values[2]
            
            check_sql = """
            SELECT 1 FROM public.perdep 
            WHERE subno = %s AND serno = %s AND depno = %s
            """
            cursor.execute(check_sql, (subno, serno, depno))
            
            if cursor.fetchone():
                delete_sql = """
                DELETE FROM public.perdep 
                WHERE subno = %s AND serno = %s AND depno = %s
                """
                cursor.execute(delete_sql, (subno, serno, depno))
                delete_count += 1
            
            insert_sql = """
            INSERT INTO public.perdep 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            cursor.execute(insert_sql, values)
            insert_count += 1
        
        conn.commit()
        cursor.close()
        conn.close()
        
        print("\n匯入完成！")
        print(f"成功插入 {insert_count} 筆資料。")
        print(f"刪除 {delete_count} 筆舊資料。")

    except FileNotFoundError:
        print(f"錯誤：找不到檔案 '{file_path}'。請確認路徑是否正確。")
    except psycopg2.Error as e:
        print(f"資料庫連線或操作發生錯誤：{e}")
        if 'conn' in locals() and conn:
            conn.rollback()

if __name__ == '__main__':
    db_config = {
        'host': 'ktmesdb.kenda.com.tw',
        'dbname': 'kterp',
        'user': 'oak',
        'password': 'Kenda_202508'
    }
    
    file_path = 'perdep.txt'
    
    import_perdep_data(db_config, file_path)
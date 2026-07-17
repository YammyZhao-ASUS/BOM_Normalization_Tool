from bom_reader import BOMReader

def main():
    # 初始化讀取器
    reader = BOMReader()
    
    # 讀取 Excel 資料
    df = reader.load()
    
    # 印出前幾筆資料確認成功
    print("\n🎉 成功讀取 Excel！以下是前 5 筆資料：")
    print(df.head())

if __name__ == "__main__":
    main()
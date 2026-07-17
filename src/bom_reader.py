import os
import pandas as pd

class BOMReader:
    def __init__(self, filename=None):
        if filename is None:
            # 自動定位到這個專案下的 input/bom.xlsx 絕對路徑
            current_dir = os.path.dirname(os.path.abspath(__file__))
            self.filename = os.path.normpath(os.path.join(current_dir, "..", "input", "bom.xlsx"))
        else:
            self.filename = filename
        
        # 這裡會印出程式「實際上」到底去哪裡找檔案，方便我們除錯！
        print(f"🔍 正在嘗試讀取 Excel 檔案，實際路徑為：{self.filename}")

    def load(self):
        # 檢查檔案到底存不存在
        if not os.path.exists(self.filename):
            raise FileNotFoundError(f"❌ 找不到 Excel 檔案！請確認此路徑下有 bom.xlsx：{self.filename}")
            
        df = pd.read_excel(self.filename)
        return df
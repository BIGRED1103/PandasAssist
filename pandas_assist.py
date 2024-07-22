# pandas 使用助手，将较为常用的功能集成到此文件中
import os

import pandas as pd

# # pandas 安装
# # 1. pip install pandas
# # 2. pip install openpyxl
# # 3. pip install jinja2 # DataFrame.style的依赖


def merge(df1, df2, on, how="outer",ignore_index=True):
    """ 
        合并两个DataFrame
        on: 指定对齐的索引列
        how: "left", "right", "inner", "outer"
        ignore_index：True 返回值不带索引列，False：返回值有索引列
        返回值：
            返回合并后的DataFrame
    """
    df = pd.merge(df1, df2, on=on, how=how)
    return df.reset_index() if ignore_index else df
    
def concat(dfs, axis=0, join="outer", ignore_index=True, keys=None):
    """ 
        拼接多个DataFrame
        axis: 0 垂直堆叠，1水平堆叠
        join: 控制如何处理索引。'outer'表示使用所有对象的联合索引，'inner'表示使用交集索引。
        ignore_index：True 返回值不带索引列，False：返回值有索引列
        返回值：
            返回拼接后的DataFrame
    """
    return pd.concat(dfs, axis=axis, join=join, ignore_index=ignore_index, keys=keys)

def iter_rows(df):
    # 遍历dataframe的行
    for idx, row in df.iterrows():
        yield idx, row
        
def iter_column(df):
    # 遍历dataframe的列
    for column in df.columns:
        yield column, df[column]
        
def clip(df, x1, y1, x2, y2):
    # 裁剪df
    return df.loc[x1:x2, y1:y2]
    
def apply_color(val):
    print(val.name)
    print(val.name in ["name", "score"])
    return ["background-color:#F9C38A"]*len(val) if val.name in ["name", "score"] else[""]*len(val)
    

class PandasExcelAssist():
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.dfs = {}

    def get_sheet_names(self):
        return pd.ExcelFile(self.excel_path).sheet_names
        
    def read_all_sheet(self):
        sheet_names = self.get_sheet_names()
        for sheet_name in sheet_names:
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            self.dfs[sheet_name] = df

    # 读取文件
    def read_excel(self, sheet_name=None):
        # 读取excel, 将页名和dataframe保存在dfs中
        # sheet_name 为None：读取第一个sheet
        # sheet_name 为页名： 读取对应页名的sheet
        # sheet_name 为页名列表：读取列表中所有的sheet
        if not sheet_name:
            sheet_name = self.get_sheet_names()[0:1]
        if not isinstance(sheet_name, list):
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            self.dfs[sheet_name] = df
            return
        for sname in sheet_name:
            df = pd.read_excel(self.excel_path, sheet_name=sname)
            self.dfs[sname] = df
        
    def write_excel(self, outpath):
        with pd.write_excel(outpath, index=False) as writer:
            for sheet_name, df in self.dfs:
                pd.to_excel(writer, df, sheet_name=sheet_name)
        
    def get_dataframe(self, sheet_name):
        return self.dfs[sheet_name]
        
    def print(self):
        for k,v in self.dfs.items():
            print("&"*8, k, "&"*    8,)
            print(v)
    
if __name__ == "__main__":
    root_dir = r"D:\A_lihong\auto_office\data"
    filename = "data1.xlsx"
    out_filename = "out.xlsx"
    file_path = os.path.join(root_dir, filename)
    pea = PandasExcelAssist(file_path)
    pea.read_excel()
    df = pea.get_dataframe(pea.get_sheet_names()[0])

    # df1 = df.style.apply(apply_color)
    # df1 = df1.set_properties(**{"border":"1px solid black"})
    
    df1 = df.style.set_properties(border="1px solid black")
    df1 =df1.set_properties(subset=["id", "score"], **{"background-color":"#F9C38A"})
    df1.to_excel(os.path.join(root_dir, out_filename), engine="openpyxl")
    
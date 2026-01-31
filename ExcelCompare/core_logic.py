import pandas as pd

class ExcelComparator:
    PART_KEYWORDS = ['part']
    QTY_KEYWORDS = ['qty', 'pcs in ctn']
    
    def read_excel_file(self, file_path):
        try:
            df = pd.read_excel(file_path, header=None, dtype=str).fillna('')

            for i in range(len(df)): # 遍历每行
                row = df.iloc[i].astype(str).str.lower()
                if any(p in ' '.join(row) for p in self.PART_KEYWORDS) and \
                   any(q in ' '.join(row) for q in self.QTY_KEYWORDS):  # 如果这一行同时出现件号和数量关键词
                    part_col = next(c for c in range(len(df.columns)) 
                                  if any(p in row[c].lower() for p in self.PART_KEYWORDS))
                    qty_col = next(c for c in range(len(df.columns)) 
                                 if any(q in row[c].lower() for q in self.QTY_KEYWORDS))
                    if part_col == qty_col:
                        continue

                    result = {}
                    for j in range(i+1, len(df)):
                        part = df.iat[j, part_col].strip()
                        qty = pd.to_numeric(df.iat[j, qty_col], errors='coerce')
                        if part and pd.notna(qty) and qty:
                            result[part] = result.get(part, 0) + int(qty)
                    return result
            
            raise ValueError("未找到件号或数量关键词！请修改 keywords.txt 中的关键词（用英文逗号分割）并重启程序")
        except Exception as e:
            raise Exception(f"读取失败: {str(e)}")
    
    def compare(self, file1, file2):
        d1 = self.read_excel_file(file1)
        d2 = self.read_excel_file(file2)
        
        result = {}
        result["前者缺失："] = {k: d2[k] for k in d2 if k not in d1}
        result["后者缺失："] = {k: d1[k] for k in d1 if k not in d2}
        result["数量差异："] = {k: d1[k]-d2[k] for k in d1 if k in d2 and d1[k] != d2[k]}
        
        return result

if __name__ == "__main__":
    comparator = ExcelComparator()
    diff = comparator.compare('1.xlsx', '2.xlsx')

    for desc, dic in diff.items():
        print(desc)
        for i in dic:
            print(f"{i}: {dic[i]}")
        print()

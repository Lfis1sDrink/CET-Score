import requests
import pandas as pd
from urllib.parse import quote
import time
import openpyxl

# 配置信息
EXCEL_PATH = "students.xlsx"
OUTPUT_PATH = "results.xlsx"
BASE_URL = "https://cachecloud.neea.cn/latest/results/cet"

def build_headers(name):
    return {
        "authority": "cachecloud.neea.cn",
        "accept": "*/*",
        "accept-encoding": "gzip, deflate, br, zstd",
        "accept-language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "origin": "https://cjcx.neea.edu.cn",
        "priority": "u=1, i",
        "referer": "https://cjcx.neea.edu.cn/",
        "sec-ch-ua": '"Not(A:Brand";v="99", "Google Chrome";v="133", "Chromium";v="133"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "cross-site",
        "sec-fetch-storage-access": "active",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36"
    }

def query_score(session, name, id_card, kmcode):
    try:
        # URL编码姓名
        encoded_name = quote(name,safe='')
        # print(encoded_name)
        # print()
        params = {
            "km": kmcode,  # 科目代码，根据实际情况修改
            "xm": name,
            "no": id_card,
            "source": "pc"
        }
        # print(params)
        # print()
        # 发送请求 
        response = session.get(BASE_URL, params=params, headers=build_headers(name))
        response.raise_for_status()
        # print(response.url)
        # print()
        # 解析响应（根据实际返回格式调整）
        result = response.json()  # 假设返回JSON格式
        # print(result)
        return result.get('score', 'N/A')
    
    except Exception as e:
        print(f"查询失败：{name} - {str(e)}")
        return "查询失败"

def main():
    # 读取Excel数据
    df = pd.read_excel(EXCEL_PATH)
    
    # 验证必要列是否存在
    required_columns = ['姓名', '身份证号']
    if not all(col in df.columns for col in required_columns):
        raise ValueError("Excel中必须包含'姓名'和'身份证号'列")

    # 初始化Session保持连接
    with requests.Session() as session:
        cet4results = []
        cet6results = []
        for index, row in df.iterrows():
            cet6score = query_score(session, row['姓名'], row['身份证号'],2)
            cet4score = query_score(session, row['姓名'], row['身份证号'],1)

            cet6results.append(cet6score)
            cet4results.append(cet4score)
            # 添加延迟防止频繁请求（建议3-5秒）
            time.sleep(3)
            
            # 实时显示进度
            print(f"进度：{index+1}/{len(df)} - {row['姓名']}：CET-6:{cet6score}  CET-4:{cet4score}")

    # 保存结果
    df['CET6成绩'] = cet6results
    df['CET4成绩'] = cet4results

    df.to_excel(OUTPUT_PATH, index=False)
    print(f"查询完成，结果已保存至：{OUTPUT_PATH}")

if __name__ == "__main__":
    main()
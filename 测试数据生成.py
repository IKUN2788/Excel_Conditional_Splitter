import pandas as pd
import random

def generate_data():
    # Create dummy data
    names = [f"学生{i}" for i in range(1, 51)]
    scores = [random.randint(40, 100) for _ in range(50)]
    classes = [random.choice(['一班', '二班', '三班']) for _ in range(50)]
    ids = [f"2023{str(i).zfill(3)}" for i in range(1, 51)] # Regex friendly: ^2023\d{3}$
    comments = [
        random.choice([
            "表现优秀，继续保持", 
            "基础扎实，但需细心", 
            "需要加强练习", 
            "进步很大", 
            "缺勤较多"
        ]) for _ in range(50)
    ]

    data = {
        "姓名": names,
        "分数": scores,
        "班级": classes,
        "学号": ids,
        "评语": comments
    }

    df = pd.DataFrame(data)
    
    output_file = "测试数据.xlsx"
    df.to_excel(output_file, index=False)
    print(f"已生成测试文件: {output_file}")

if __name__ == "__main__":
    generate_data()

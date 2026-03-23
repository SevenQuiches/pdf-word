import os
from pdf2docx import Converter

# 测试基本功能
def test_basic():
    print("Testing pdf2docx library...")
    
    # 测试导入是否成功
    print("pdf2docx 导入成功")
    
    # 创建一个测试用的PDF文件路径（假设存在的话会测试转换）
    test_pdf = "test.pdf"
    if os.path.exists(test_pdf):
        print(f"找到测试文件: {test_pdf}")
        try:
            cv = Converter(test_pdf)
            print("Converter 对象创建成功")
            cv.close()
        except Exception as e:
            print(f"Converter 测试: {e}")
    else:
        print("没有找到test.pdf文件，跳过转换测试")
    
    print("\n测试完成!")

if __name__ == "__main__":
    test_basic()

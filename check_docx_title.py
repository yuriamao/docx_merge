import os  
from docx import Document

# https://python-docx.readthedocs.io/en/latest/
# pip install python-docx


output_folder = "data/output"

if not os.path.exists(output_folder):  
    os.makedirs(output_folder)  
    print(f"Path {output_folder} created.")  
else:  
    print(f"Path {output_folder} already exists.")  

document=Document()

# 获取文件列表并排序  
file_list = os.listdir("data/05 产业上中下游价格指数周度报告/产业上中下游价格指数报告_49篇")  
file_list.sort()
print(file_list)
# 创建一个新文件，如果文件已经存在则会被覆盖  

with open("output.txt", "w") as file:  
    file.write(f"输出结果：\n")  
    for i, filename in enumerate(file_list):  
        if filename.endswith(".docx"):  
            doc_path = os.path.join("data/05 产业上中下游价格指数周度报告/产业上中下游价格指数报告_49篇", filename)  
            doc = Document(doc_path)  
            first_paragraph = doc.paragraphs[0].text  
            second_paragraph = doc.paragraphs[1].text  
            first_paragraph = f'{first_paragraph} 第（{i+1}）期'  
            print(filename, first_paragraph, second_paragraph)  
            doc.paragraphs[0].text = first_paragraph  
            
            output_filename = os.path.join(output_folder, filename)  
            doc.save(output_filename)  
            with open('output.txt', 'a+') as f:  
                f.write(f"{filename}: {first_paragraph}")  
                f.write(f"{second_paragraph}\n")
            print(f"文档已修改并保存到：{output_filename}")
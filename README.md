# check_docx_tilte

## 功能

- [X] 修改每个文档的第一段话后+第x期（会破坏原有的样式）
- [X] 输出每个word文档第一段话和第二段话的在/output.txt
- [ ] 把正文的第一句话改为标题级别

## 用法

1. 在当前环境下安装 `python-docx`库：pip install python-docx
2. 在根目录下运行：
   1. check_docx_title.py 修改第一段话名称
   2. merge.py 合并docx到同一个文件

# merge

实现自动合并word的python脚本文件

- 保留所有文档固有的样式
- 保留文档的顺序

## 使用方法

**【注意】** 只能合并 `.docx`文件，若需要合并的文档中存在 `.doc`文件，需先手动将其转换为 `.docx`文件才能使用此脚本

- 在该项目的根目录下创建 `files`文件夹，将需要合并的 `.docx`文件放进去
- 运行 `merge.py`，即可得到合并结果文件 `merge_result.docx`

## 特性说明

### 文档分页

每个文档自动从新一页开始，两个文档之间会插入一个分页符

### 文档合并顺序

文档的合并顺序与在资源管理器中的排序一致，因此对排序有强要求的需求，最好对文档进行编号

# 鸣谢

参考引用：https://github.com/FutureXZC/auto-merge-docx（代码merge.py）

常见文档批量操作：https://zhuanlan.zhihu.com/p/323680114

mac 安装brew：https://www.cnblogs.com/liyihua/p/12753163.html

doc转docx：

https://zhuanlan.zhihu.com/p/649925993

https://www.cnblogs.com/bubblebeee/p/17096397.html

https://blog.csdn.net/S1mpleboy6/article/details/132190913

https://zhuanlan.zhihu.com/p/561923128

https://www.wenjiangs.com/article/r05nnmmsj10d.html

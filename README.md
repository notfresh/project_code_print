# 说明
打印源代码到word文档里面, 方便纸质阅读. 简易树形图, 压缩代码行间距, 尽量节省纸张

# 代码文件说明
- test_docx  
这是一个写docx测试用的目录. 与本项目可用性无关. 喜欢玩docx工具包的同学可以自己去玩一下.  
- util.py  
这是这个项目的主要核心方法.  
- run.py  
这是这个项目的启动方法.  
- project_code_print.py  
这个文件是我刚刚新加入进来的, 可以以一个单源代码文件的形式执行代码打印.  

# 使用方法
0. 首先你的电脑上要装有python3.6, pip等工具
1. pip install python-docx==0.8.10
2. 进入要打印的目录下面, 把project_code_print.py拷贝进去
3. 运行 python project_code_print.py 即可, 会在当前目录下生成一个 `ode_print.docx` 的文件. 
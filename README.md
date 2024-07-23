# W2Md
一、项目地址
https://github.com/xxhanzo/Word2Markdown
二、项目结构
- `Word2Markdown` 文件夹
  - `W2Md_1` 文件夹
    - `W2M_upload.py`: 通过上传来转换md文件
    - `W2M_path.py` : 通过修改路径来转换md文件
  - `data` 文件夹
    - 包含的是测试数据。
其中只需要运行`W2M_upload.py`通过上传来转换md文件即可，方便
三、效果展示
1. 安装`requirements.txt`
pip install -r requirements.txt
2. 运行
可以一次性上传多个文件，点击打开即可
运行结果：

![image](https://github.com/user-attachments/assets/ac47aceb-a958-4ba1-9a2f-ac3b1237ac77)

生成的结果在`Word2Markdown\W2Md_1\generate_data` 文件夹下：
![image](https://github.com/user-attachments/assets/d4500e2b-3776-4150-aa5c-27d652d8aeef)

3. 功能及限制条件
1. 可以识别文本，图片，表格并在正确的位置
  1. 以下是测试表格的效果图
![image](https://github.com/user-attachments/assets/799ac0ab-1213-4e3e-b27d-d764c32085cb)

2. 限制条件
  1. 包含第一页封面
  2. 二级标题为类似：”1范围“，”2规范性引用文件“
  3. 三级标题为类似：”4.1“，”5.2“
  4. 依次类推，四级标题：“4.1.1”，五级标题：“5.2.1.2”

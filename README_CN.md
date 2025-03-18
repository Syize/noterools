# 创建可跳转的Zotero引用

Zotero插入的引用和参考文献表不能像正常的交叉引用一样跳转，这个Python脚本可以对Zotero的引用进行处理，设置超链接实现跳转功能

## 原理

- 扫描参考目录表中的所有条目，添加相应书签
- 扫描正文中的所有引用，对年份设置跳转超链接
- 这个脚本也能同时设置引用的字体颜色，参考文献表中出版物或出版商的斜体，以及去除超链接的下划线

## 注意

- **该脚本仅能在Windows环境下运行**
- **顺序引用格式没有被测试过**

## 如何使用

1. 克隆本仓库
2. 安装以下依赖
   - pywin32
   - pyzotero
   - rich
3. 修改`main.py`：
   - `zotero_id`和`zotero_api_key`被用于与Zotero通信，获取方式请查看[pyzotero](https://pyzotero.readthedocs.io/en/latest/index.html)的手册。
   - `word_file_path`是你的Word文档的绝对路径。
   - `new_file_path`是新保存的文档的绝对路径。
   - `isNumbered`是布尔值，代表你的引用格式是否为顺序引用。
   - `setColor`是代表字体颜色的整数值。你可以查阅[VBA的文档](https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor)获得其他颜色对应的值。默认是蓝色。
   - `noUnderLine`是布尔值，代表是否去除超链接的下划线。
   - `set_container_title_italic`是布尔值，代表是否将出版物或出版商设置为斜体。目前Zotero插入的参考文献表中不能将中文出版物或出版商正确设置为斜体，这个问题同样存在于韩文或日本中。**这个特性仅对中文测试过。
4. 运行`main.py`

# 创建可跳转的Zotero引用

Zotero插入的引用和参考文献表不能像正常的交叉引用一样跳转，这个Python脚本可以对Zotero的引用进行处理，设置超链接实现跳转功能

## 原理

- 扫描参考目录表中的所有条目，添加相应书签
- 扫描正文中的所有引用，对年份设置跳转超链接

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
   - `zotero_id`和`zotero_api_key`被用于与Zotero通信，获取方式请查看[pyzotero](https://pyzotero.readthedocs.io/en/latest/index.html)的手册
   - `word_file_path`是你的Word文档的绝对路径
   - `new_file_path`是新保存的文档的绝对路径
4. 运行`main.py`

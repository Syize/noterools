# 创建可跳转的Zotero引用

Zotero插入的引用和参考文献表不能像正常的交叉引用一样跳转，这个Python脚本可以对Zotero的引用进行处理，设置超链接实现跳转功能

## 原理

脚本原理参考了[gwyn-hopkins](https://forums.zotero.org/discussion/comment/418013/#Comment_418013
)的实现，将其移植到Python中并做了一些改进。根据我的测试，对`(作者, 年份)`格式的引用修改效果良好。

脚本的运作原理如下：

- 扫描参考目录表中的所有条目，添加相应书签
- 扫描正文中的所有引用，对年份设置跳转超链接

## 使用

**该脚本仅对[GB/T 7714—2015（著者-出版年，双语，姓名不大写，无 URL、DOI，全角括号）](https://zotero-chinese.com/styles/gb-t-7714-2015-author-date-bilingual-no-uppercase-no-url-doi-fullwidth-parentheses/)做过测试，并且条目的作者命名方式，以及语言设置均会影响到脚本的运行，如果出现错误，你可以对脚本自行修改**

**该脚本仅能在Windows环境下运行**

### 安装依赖库

```bash
pip install pywin32 rich
```

### 运行

1. 修改代码中的输入文件路径`word_file_path`和保存文件路径`new_file_path`，最好设置为绝对路径
2. 运行脚本

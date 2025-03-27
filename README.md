# Noterools: Not just Zotero Tools

<p align="center">中文文档 | <a href="README_EN.md">English</a></p>

一开始我只是想依照[gwyn-hopkins](https://forums.zotero.org/discussion/comment/418013/#Comment_418013)的代码写一份相应的Python实现，用于为Zotero的引用添加可跳转的超链接。但是随着论文的修改，我发现需要对论文的格式做越来越多的设置，代码实现的功能也越来越多。经过大量的重构以后，noterools诞生。

## 这是什么?

目前noterools包含以下功能：

- 为Zotero参考文献表中的每个文献创建书签
- 为Zotero的引用设置跳转到相应文献的超链接，并设置超链接是否带下划线
- 为Zotero的引用设置字体颜色
- 将Zotero的参考文献表中，不能被正确设置为斜体的期刊名称和出版商设置为斜体
- 为正文中的交叉引用设置字体颜色和粗细

## 效果图

![引用和参考文献表设置](./pics/noterools1.png)

![交叉引用设置](./pics/noterools2.png)

## 注意

- **该脚本仅能在Windows环境下运行**
- **为顺序引用格式添加超链接的功能正常，但是暂时不能为期刊名称和出版商设置斜体**

## 如何使用

1. 使用pip安装noterools
```bash
pip install noterools
```
2. 创建一个Python脚本并运行。以下是一个简单的示例
```python
from noterools import Word, add_citation_cross_ref_hook, add_cross_ref_style_hook

if __name__ == '__main__':
    word_file_path = r"E:\Documents\Word\test.docx"
    new_file_path = r"E:\Documents\Word\test_new.docx"

    with Word(word_file_path, save_path=new_file_path) as word:
        # 为Zotero的citation添加超链接
        add_citation_cross_ref_hook(word, is_numbered=False, color=16711680, no_under_line=True, set_container_title_italic=True)
        # 为正文中以Figure开头的交叉引用设置字体颜色和粗体
        add_cross_ref_style_hook(word, color=16711680, bold=True, key_word=["Figure"])
        # 执行操作
        word.perform()
```

# Noterools: Not just Zotero Tools

<p align="center"><a href="README.md">中文文档</a> | English</p>

At the beginning, I just wanted to write a Python implementation based on [gwyn-hopkins](https://forums.zotero.org/discussion/comment/418013/#Comment_418013)'s code to add clickable hyperlinks to Zotero citations. However, as my paper underwent more revisions, I found myself needing to make increasingly complex formatting adjustments. Consequently, the functionality of the code expanded. After extensive refactoring, noterools was born.

## What is this?

Currently, noterools can help you do the following things:

- Create bookmarks for each reference in the Zotero bibliography.
- Set hyperlinks for Zotero citations that navigate to the corresponding references and adjust whether the hyperlinks are underlined.
- Customize the font color of Zotero citations.
- Italicize journal names and publishers in the Zotero bibliography that aren't correctly formatted as italics.
- Adjust the font color and weight for cross-references within the main text.

## Screenshots

![citation and bibliography](./pics/noterools1.png)

![cross-references](./pics/noterools2.png)

## Important Note

- **This script can only work in Windows.**
- **The function for adding hyperlinks to sequential citation formats works properly, but italicizing journal names and publishers is not yet supported.**

## How to use?

1. Install noterools via pip.
```bash
pip install noterools
```
2. Create a Python script and run it. Here is a simple example.
```python
from noterools import Word, add_citation_cross_ref_hook, add_cross_ref_style_hook

if __name__ == '__main__':
    word_file_path = r"E:\Documents\Word\test.docx"
    new_file_path = r"E:\Documents\Word\test_new.docx"

    with Word(word_file_path, save_path=new_file_path) as word:
        # Add hook to add hyperlinks to Zotero citations.
        add_citation_cross_ref_hook(word, is_numbered=False, color=16711680, no_under_line=True, set_container_title_italic=True)
        # Add hook to set the font color and bold style for cross-references starting with 'Figure' in the main text.
        add_cross_ref_style_hook(word, color=16711680, bold=True, key_word=["Figure"])
        # Perform
        word.perform()
```

from src.run import run


if __name__ == '__main__':
    zotero_api_key = ""
    zotero_id = ""
    word_file_path = r""
    new_file_path = r""
    run(word_file_path, new_file_path, zotero_id, zotero_api_key, isNumbered=False, setColor=16711680, noUnderLine=True)

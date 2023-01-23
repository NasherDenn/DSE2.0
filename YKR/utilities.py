# функция получения названий файлов для дальнейшего получения из них данных
def name_dir(name_dir_files):
    # переменная-список для дальнейшего преобразования списка списков в список строк выбранных для загрузки файлов docx
    name_dir_docx = []
    for i in name_dir_files[:-1]:
        name_dir_docx.append(i)
    return name_dir_docx
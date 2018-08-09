# coding:gb18030
import os,zipfile,shutil
from win32com import client

def doc2docx(doc_name,docx_name):
    """
    :doc转docx
    """
    # 首先将doc转换成docx
    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_name)
    #使用参数16表示将doc转换成docx
    doc.SaveAs(docx_name,16)
    doc.Close()
    word.Quit()
def extract_img(docx_name,docdir):
    docname = docx_name.split(".d")  # 以“.”做成列表形式
    if docx_name.endswith(".doc"):
        doc_name = "%s/%s.doc" % (docdir,docname[0])
        docx_name = "%s/%s.docx" % (docdir,docname[0])
        doc2docx(doc_name,docx_name)
    os.rename(docx_name, "%s.ZIP" % docname[0])  # 重命名为ZIP格式
    f = zipfile.ZipFile("%s.ZIP" % docname[0], 'r')
    for file in f.namelist():
        if "word" in file:
            f.extract(file)  # 将压缩包里的word文件夹解压出来
    f.close()
    oldimagedir = r"%s/word/media" % docdir  # 定义图片文件夹
    shutil.copytree(oldimagedir, "%s/%s" % (docdir, docname[0]))  # 拷贝到新目录,名称为word文件的名字
    os.rename("%s.ZIP" % docname[0], "%s.docx" % docname[0])  # 将ZIP名字还原为DOCX
    shutil.rmtree("%s/word" % docdir)  # 删除word文件夹

def getimage(docdir):
    os.chdir(docdir)
    dirlist = os.listdir(docdir)
    for docx_name in dirlist:
        print('正在保存文件<%s>中的图片.....'% docx_name)
        if docx_name.endswith(".doc"):
            docname1 = docx_name.split()  # 以“.”做成列表形式
            docx_name_nospace = ''.join(docname1)
            os.rename(docx_name,docx_name_nospace)  # 重命名删除空格
            extract_img(docx_name_nospace,docdir)
        else:
            extract_img(docx_name, docdir)
if __name__=="__main__":
    path = r"F:/Users/18435/PycharmProjects/learn/7.27动漫"
    getimage(path)
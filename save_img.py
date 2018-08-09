# coding:gb18030
import os,zipfile,shutil
from win32com import client

def doc2docx(doc_name,docx_name):
    """
    :docתdocx
    """
    # ���Ƚ�docת����docx
    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_name)
    #ʹ�ò���16��ʾ��docת����docx
    doc.SaveAs(docx_name,16)
    doc.Close()
    word.Quit()
def extract_img(docx_name,docdir):
    docname = docx_name.split(".d")  # �ԡ�.�������б���ʽ
    if docx_name.endswith(".doc"):
        doc_name = "%s/%s.doc" % (docdir,docname[0])
        docx_name = "%s/%s.docx" % (docdir,docname[0])
        doc2docx(doc_name,docx_name)
    os.rename(docx_name, "%s.ZIP" % docname[0])  # ������ΪZIP��ʽ
    f = zipfile.ZipFile("%s.ZIP" % docname[0], 'r')
    for file in f.namelist():
        if "word" in file:
            f.extract(file)  # ��ѹ�������word�ļ��н�ѹ����
    f.close()
    oldimagedir = r"%s/word/media" % docdir  # ����ͼƬ�ļ���
    shutil.copytree(oldimagedir, "%s/%s" % (docdir, docname[0]))  # ��������Ŀ¼,����Ϊword�ļ�������
    os.rename("%s.ZIP" % docname[0], "%s.docx" % docname[0])  # ��ZIP���ֻ�ԭΪDOCX
    shutil.rmtree("%s/word" % docdir)  # ɾ��word�ļ���

def getimage(docdir):
    os.chdir(docdir)
    dirlist = os.listdir(docdir)
    for docx_name in dirlist:
        print('���ڱ����ļ�<%s>�е�ͼƬ.....'% docx_name)
        if docx_name.endswith(".doc"):
            docname1 = docx_name.split()  # �ԡ�.�������б���ʽ
            docx_name_nospace = ''.join(docname1)
            os.rename(docx_name,docx_name_nospace)  # ������ɾ���ո�
            extract_img(docx_name_nospace,docdir)
        else:
            extract_img(docx_name, docdir)
if __name__=="__main__":
    path = r"F:/Users/18435/PycharmProjects/learn/7.27����"
    getimage(path)
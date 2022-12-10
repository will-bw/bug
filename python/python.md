# 日常

## 2022年12月10日21:13:43

> 从网上找的批量转换PDF的脚本，运行报错 module 'win32com.gen_py.91493440-5A91-11CF-8700-00AA0060263Bx0x2x12' has no attribute 'MinorVersion'

~~~python
import os
import glob
from win32com.client import gencache
 
 
def get_file_path():
    """
    获得当前文件夹下的所有的.ppt和.pptx文件
    """
    file_path = os.path.split(os.path.abspath(__file__))[0]
    pp_files = glob.glob(os.path.join(file_path, "*.ppt*"))
    return file_path, pp_files
 
def ppt_to_pdf(filename, results):
    '''
    ppt 和 pptx 文件转换
    '''
    name = os.path.basename(filename).split('.')[0] + '.pdf'
    exportfile = os.path.join(results, name)
    if os.path.isfile(exportfile):
        print(name, "已经转化了")
        return
    p = gencache.EnsureDispatch("PowerPoint.Application")
    try:
        ppt = p.Presentations.Open(filename, False, False, False)
    except Exception as e:
        print(os.path.split(filename)[1], "转化失败，失败原因%s" % e)
    ppt.ExportAsFixedFormat(exportfile, 2, PrintRange=None)
    print('保存 PDF 文件：', exportfile)
    p.Quit()
 
def main():
    """
    主程序执行
    """
    file_path, pp_files = get_file_path()
    results = os.path.join(file_path, "results")
    if not os.path.exists(results):
        os.mkdir(os.path.join(results))
    for _ in pp_files:
        ppt_to_pdf(_, results)
 
if __name__ == "__main__":
    main()
~~~



![image-20221210211416872](https://gitee.com/wang-bangwen/image/raw/master/img/image-20221210211416872.png)


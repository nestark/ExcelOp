import os
import pandas

def getFileName(path):
    ''' 获取指定目录下的所有指定后缀的文件名 '''
    f_list = os.listdir(path)
    sf_list=[]
    # print f_list
    for i in f_list:
        if os.path.splitext(i)[1] == '.xls':
            xls2xlsx(path,i)
        # os.path.splitext():分离文件名与扩展名
        elif os.path.splitext(i)[1] == '.xlsx':
            sf_list.append(i)
            print(i)
    return sf_list

def xls2xlsx(f_path,f_name):
    name=f_name.split('.')
    file=pandas.read_excel(f_path+f_name)
    file.to_excel(f_path+name[0]+'.xlsx',index=False)


if __name__ == '__main__':

    path ='./'
    getFileName(path)


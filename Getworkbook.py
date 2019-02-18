import os


def getFileName(path):
    ''' 获取指定目录下的所有指定后缀的文件名 '''
    f_list = os.listdir(path)
    sf_list=[]
    # print f_list
    for i in f_list:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(i)[1] == '.xlsx':
            print(i)
            sf_list.append(i)
    return sf_list

if __name__ == '__main__':

    path ='.'
    getFileName(path)

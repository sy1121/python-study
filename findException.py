'''
Created on 2017-5-9

@author: sy
'''
#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
def readFile(filename):
    count=0
    flag=0
    result=""
    list=[]
    try:
        fopen = open(filename, 'r') 
        for eachLine in fopen:
            if ( flag == 1 ) :
                if eachLine.startswith("//"):
                    result+=eachLine+"\n"
                else:
                    flag=0
                    list.append(result)
                  #  print result
                    result=""
            if ("CRASH:" in eachLine):
                result+=str(count)+". \n"
                result+=eachLine+"\n"
                flag=1
                count=count+1
    except IOError,ValueError:
        print "Error: sdfasd"
    print "listlen",len(list)
    return list
def findException(filepath):
    pathDir =  os.listdir(filepath)
    exception="D:\FDroid-38\\exception"
    for allDir in pathDir:
        child = os.path.join('%s\%s' % (filepath, allDir))
       # os.mknod('%s\%s' % (child, "exception-r.txt")) 
       # print child.decode('gbk') 
        if os.path.isdir(child):
            l=readFile('%s\%s' % (child, "info.txt"))
            if ( len(l) == 0 ): continue
            exce= os.path.join('%s\%s' % (exception, allDir+"-exception-r.txt"))
           # if os.path.exists(exce): os.remove(exce)
            file_object = open(exce, 'a')
            file_object.write(l[0])
            file_object.close()

if __name__ == '__main__':
    filePathC = "D:\FDroid-38\\Monkey-all"
    findException(filePathC)

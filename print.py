#!/usr/bin/python
import paramiko
import os
import win32api
import win32print
import time
import threading
import shutil
import getpass

# 服务器地址
hostname = ''
# 账号
username = ''
# 端口
port = 22
# 本地地址
local_dir = os.getcwd()+'\\data\\'
# 远程地址
remote_dir = '/home/wwwroot/erp/temp/excel/'

pathHistory = os.getcwd()+"\\history\\"

def getRemote(password):
    try:
        t = paramiko.Transport((hostname, port))
        t.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(t)
        files = sftp.listdir(remote_dir)
        for f in files:
            sftp.get(os.path.join(remote_dir, f), os.path.join(local_dir, f))
            sftp.remove(os.path.join(remote_dir, f))
        t.close()
    except Exception as err:
        print(err)
    # t = threading.Timer(5, getRemote)
    # t.start()

def printExcel():
  dirs = os.listdir(local_dir)
  for fileName in dirs:
    if os.path.splitext(fileName)[1] == ".xlsx":
      print('打印文件'+fileName)
      win32api.ShellExecute(
        0,
        "print",
        fileName,
        '/d:"%s"' % win32print.GetDefaultPrinter(),
        local_dir,
        0
      )
      time.sleep(10)  # 获取指定路径下的文件
      moveFile(fileName)

  # #10s执行一次
  # t = threading.Timer(10, printExcel)
  # t.start()

def moveFile(fileName):
  today = time.strftime("%Y%m%d", time.localtime())
  if (os.path.exists(pathHistory+today)):
    shutil.move(local_dir+fileName, pathHistory+today+'\\'+fileName)
  else:
    os.makedirs(pathHistory+today)
    shutil.move(local_dir+fileName, pathHistory+today+'\\'+fileName)

def index(password):
    getRemote(password)
    printExcel()
    t = threading.Timer(1, index(password))
    t.start()

def connetRemote(password):
    try:
        t = paramiko.Transport((hostname, port))
        t.connect(username=username, password=password)
        t.close()
        return True
    except Exception as err:
        return False

if __name__ == "__main__":
  print('自动打印程序已启动.....................')
  print('请勿关闭该窗口.....................')
  while True:
      password = getpass.getpass("请输入远程服务器密码：")
      if connetRemote(password):
          print('连接成功.....................')
          if (os.path.exists(local_dir)):
              index(password)
          else:
              os.makedirs(local_dir)
              index(password)
      else:
          print('密码错误，请重新输入')


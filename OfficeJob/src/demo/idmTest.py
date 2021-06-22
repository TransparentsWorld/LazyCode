from subprocess import call

IDM = r'C:\Application\SystemTool\Internet Download Manager\IDMan.exe'
DownUrl = r'https://ss0.bdstatic.com/70cFvHSh_Q1YnxGkpoWK1HF6hhy/it/u=1818910735,2722708107&fm=26&gp=0.jpg'
DownPath = 'D:\\'
OutPutFileName = 'test.jpg'
call([IDM, '/d',DownUrl, '/p',DownPath, '/f', OutPutFileName, '/n', '/a'])
print("download end")
# 度盘短信备份导出

### 1. 下载可执行文件或自行编译

[https://github.com/rroy233/smsBackupExport/releases
](https://github.com/rroy233/smsBackupExport/releases)

### 2. 获取cookie

 - 前往 https://duanxin.baidu.com/ ，登录
 - 按`F12`打开开发者工具
 - console内输入`var cookie=document.cookie;var ask=confirm('Cookie:'+cookie+'\n\n是否复制到剪贴板?');if(ask==true){copy(cookie);msg=cookie}else{msg='Cancel'}`
 - 复制后粘贴到与可执行文件同目录下的`cookie.txt`中。

### 3. 运行程序

导出文件名为`out.xlsx`
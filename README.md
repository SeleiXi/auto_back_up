# auto_back_up
開機自動備份重要文件到github + 本地的其他硬盤
將autoBackUp.vbs放入C盤根目錄（C:\）
將其他兩個文件放入C:\Users\【user_name】\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup    （win+r後輸入shell:startup 即可進入，放入後文件會在開機時自動執行）
請自行在文件中填入本地的文件名以及github的倉庫名等（在本地也需要配置好git環境），才能執行程序

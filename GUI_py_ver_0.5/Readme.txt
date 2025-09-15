最後更新時間: 2025/9/15

使用說明:
    1. 配置程式執行環境(一次性)
    2. 執行程式
    3. 數據儲存

1.<<< 配置程式執行環境 >>>

    >>> 下載並安裝 Tesseract ORC：https://github.com/UB-Mannheim/tesseract/wiki 

        安裝完成後，將路徑加入電腦環境變數： 

        控制台 >>> 編輯系統環境變數 >>> 環境變數 >>> 選擇變數"Path" >>> 新增 >>> 貼上"C:\Users\Username\AppData\Local\Programs\Tesseract-OCR\" 

        >>> 注意: Username 要改成自己電腦的帳戶名稱，如果 Tesseract-OCR 不是安裝在Tesseract ORC安裝執行檔的預設的位置，路徑要改成 tesseract.exe 所在的資料夾


    >>> 下載並安裝 python：https://www.python.org/downloads/
        
        安裝完成後，將路徑加入電腦環境變數
            1: "C:\Users\Username\AppData\Local\Programs\Python\Python313\"
            2: "C:\Users\Username\AppData\Local\Programs\Python\Python313\Scripts\"
        
        >>> 注意: "Python313" 的數字會根據下載的 python 版本而變化


    安裝 python 套件( pip 指令，在 cmd 執行)： 
        pip install numpy opencv_python opencv_python_headless openpyxl Pillow PyAutoGUI pytesseract PyYAML matplotlib

    2025/9/15版本:
        (個套件版本號待補)

2.<<< 執行程式 >>>

    1.程式有兩種版本:
        1.Code_ipynb：將各個功能分開，以程式區塊執行各功能(需有ipynb環境，適合進階使用者)
        2.GUI_py：點擊 launcher.exe 使用 GUI 進行操作

    2.於 hospital_doctors.yaml 設定常用的醫院、URL、醫生名(如果沒此需求可跳過)

    3.執行 QueueNumberTracker.py，開啟 GUI

    ( GUI 說明 )

3.<<< 數據儲存 >>>

    程式會將存放數據的 Excel 檔配置在當前目錄的資料夾"data"裡(如果資料夾data不存在，會建立他它)
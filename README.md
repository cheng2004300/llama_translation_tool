一個VB6的小程式
用來配合SnipDo熒幕選詞,將各種prompt請求送到llama.cpp並呈現結果
現階段只有翻譯用途,配合科學Marker的TaiwanPro模型使用
其他llama.cpp可用的模型也可以但是沒有詳細測試

1.
安裝SnipDo
https://snipdo-app.com/


2.
安裝LlamaTranslation.pbar
一個SnipDo的腳本,可以設定一下,這裡直接設定成圈選就跳提示

3.
啟動AI模型服務在62008,將模式,Taiwanpro-reason_q4.gguf放在llama.cpp文件夾下
去 https://github.com/ggml-org/llama.cpp/releases
下載對應的llama.cpp版本,這裡選擇通用的avx2版本
llama-b4732-bin-win-avx2-x64

以命令啟動llama.cpp服務器
.\llama-server.exe -m Taiwanpro-reason_q4.gguf --host 0.0.0.0 --port 62008
	備註:
	llama-b4756-bin-win-sycl-x64.zip(Intel Gpu用)
	cudart-llama-bin-win-cu11.7-x64.zip(Nvida Gpu用)
	如果是GPU版本加上額外參數-ngl 33,加載模型到gpu顯存

4.
打開LlamaTranslation.exe
(首次啟動以管理員啟動才會自動控件依賴)
可以先按縮小會縮小到右下角任務欄

5.圈選一段外文送出查看結果

參考:
安裝教學
https://www.youtube.com/watch?v=JGeuWTPGOJY

參考專案:
https://github.com/CherryHQ/cherry-studio/
https://github.com/openai-translator/openai-translator/






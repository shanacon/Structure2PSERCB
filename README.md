# Structure2PSERCB

## 程式功能
### 此程式會抓取CXX.DAT MASS103.INP CNK1.INP三檔案中的資訊並生成北科大制定的PSERCB標準的excel檔

## 使用方式及注意事項
* 把 Structure2Excel.exe 放在和 CXX.DAT MASS103.INP CNK1.INP 同一個資料夾下，執行 Structure2Excel.exe會產生PSERCB.xlsx檔，exception發生的情況會產生OutputLogs.txt
* 注意有相同檔名的情況下會直接覆蓋
* 目前是假設CXX裡會把所有F和C的組合都會列出來，例如 CXX 檔案裏前兩個數是20和19，那下面就會有20 * 19個case
* 目前假設MASS103裡[FLOOR,  Hn,      W]這行下面的樓層數是由高到低且中間不會缺項
* 目前假設CNK1裡[$ ----- col.data --------]這行下下面幾行的C是由低到高且中間不會缺項
* 建議在CXX檔案下樓層柱編號的部分(ex:R1   , C1)樓層的部分不要以'F'結尾
* 只要最後沒有[Complete!!]就代表有exception

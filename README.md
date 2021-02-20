# excel 產生多個 worksheet 及樣式
> 請務必讀過 README ，並且修改 node 內容

## 在 `npm run serve | build` 前要去 node_modules 修改檔案內容

1. 下載本專案 `npm install` 之後，需要去 `node_modules` 修改 `xlsx-style` 資料夾的內容

2. 請到 `/node_modules/xlsx-style/dist/` 資料夾找出 `cpexcel.js` 檔案修改第 807行

```
把var cpt = require('./cpt' + 'able');
改成var cpt = cptable;
```
3. 完成後才可以順利運行  `npm run serve | build`
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

> 前言如下 使用者就應該被拿出來鞭屍
> ![使用者就應該被拿出來鞭屍](https://trello-attachments.s3.amazonaws.com/5f20e6dc8ba6238874c056e2/5f6ac457c0ce208b7f4d8a0f/df8b542cd9e56a8029ab1100c87418dc/%E8%A8%BB%E8%A7%A3_2020-10-07_145859(1).jpg)
> 完畢

找了很多方法，最後還是用廣為流傳的 `Blob.js` 和 `Export2Excel.js`，這兩個檔案是基於 Vue 和 [js-xlsx](https://github.com/SheetJS/sheetjs) 之間的補強檔案。

## npm
```
npm install file-saver xlsx script-loader xlsx-style --save
```

## 起手式匯出檔案

首先來完成簡單的 JSON 匯出成 excel 

1. 套件
```
npm install -S file-saver xlsx
npm install -S script-loader
```
2. 接著在 `src` 資料夾下產生一個 `Excel` 資料夾，產生 [Blob.js](https://github.com/Nighthree/excel_worksheets_style/blob/main/Blob.js) 和 [Export2Excel.js](https://github.com/Nighthree/excel_worksheets_style/blob/main/Export2Excel.js) 檔案，請記得把連結裡面的程式碼複製進去蛤~~不要怪我沒講~~

3. vue 檔內容
```HTML
<template>
  <div class="hello">
    <button @click="export2Excel">匯出表格</button>
  </div>
</template>

<script>
export default {
data () {
return {
  tableData: [
    {'index':'999',"nickName": "的舊時光", "name": "98491231841251"},
    {'index':'1',"nickName": "高貴", "name": "張"},
    {'index':'2',"nickName": "海aaa灰塵", "name": "小蘭"}
  ]
}
},
methods: {
export2Excel() {
  require.ensure([], () => {
    const { export_json_to_excel } = require('../Excel/Export2Excel'); // 你的 Excel 資料夾的位置
    const tHeader = ['序號', '暱稱', '姓名'];  // 設置 Excel 的表格第一行的標題
    const filterVal = ['index', 'nickName', 'name'];  // index、nickName、name 是 tableData 裡屬性質，表格會依此順續排列
    const list = this.tableData;  //把 data 裡的 tableData 存到list
    const data = this.formatJson(filterVal, list);
    export_json_to_excel(tHeader, data, '匯出的文件名稱');  // 匯出 Excel 文件名稱
  })
},
formatJson(filterVal, jsonData) {
      // 資料處理
       return jsonData.map(v => filterVal.map(j => v[j]))
     }
  }
}
<script>
```
這樣就完成了簡易的匯出，你沒看錯就是連 main.js 都不用引入什麼東西，在方法 `export2Excel()` 引入完成了

## 產生分頁

> 這邊要改的東西比較多，會稍微複雜點

### 一個分頁內容就是一個陣列
有一個要點是，上面的範例是一個檔案，一個檔案就是一個陣列內容，兩個檔案就是兩個陣列，也就是說內容會是這個樣子

```
tableData: [
  [
    {'index':'999',"nickName": "的舊時光", "name": "98491231841251"},
    {'index':'1',"nickName": "高貴", "name": "張"},
    {'index':'2',"nickName": "海aaa灰塵", "name": "小蘭"}
  ],
  [
    {'index':'999',"nickName": "的舊時光", "name": "98491231841251"},
    {'index':'1',"nickName": "高貴", "name": "張"},
    {'index':'2',"nickName": "海aaa灰塵", "name": "小蘭"}
  ]
]
```

`method` 的方法 `export2Excel` 內容要修改成
```javascript=
exportExcel() {
      require.ensure([], () => {
        // 產生多個分頁
        const { export_json_to_excel } = require("../Excel/Export2Excel");
        const tHeader = [['序號', '暱稱', '姓名'],['序號', '暱稱', '姓名']];
        const filterVal = [
          ["id", "title", "author"],
          ["id", "title", "author"],
        ];
        const list = this.tableData;
        
        let data = [];
        for (let i = 0; i < list.length; i++) {
          let arr = this.formatJson(filterVal[i], list[i]);
          data.push(arr);
        }
        // 分頁名稱
        const sheetArray = ["Sheet1", "Sheet2"];

        export_json_to_excel(tHeader, data, "导出文件名", sheetArray);
      });
}
```


### Export2Excel.js
`Export2Excel.js` 裡頭有一個 function `sheet_from_array_of_arrays` 是在處理產生的格子，裡面不會讓傳入的 data 變多或是變少

![](https://trello-attachments.s3.amazonaws.com/5f20e6dc8ba6238874c056e2/5f6ac457c0ce208b7f4d8a0f/51b64d561ae0acc27fa41c9d86277c49/%E8%A8%BB%E8%A7%A3_2020-10-07_171442.jpg)

所以在發動 `sheet_from_array_of_arrays` 之前的 function 就是在處理這個 data 的內容，這裡會動到的是 `export_json_to_excel` 這個函式

在許多分頁的範例會有類似下面這樣的方法
```javascript=
 for (var i = 0; i < th.length; i++) {
   data[i].unshift(th[i]);
 }
```
但是不知道為什麼傳到 `sheet_from_array_of_arrays` 裡的 data 會被連續插入兩次的表頭，造成輸出不正確

所以這裡直接捨棄這個方法，把表頭和 jsonData 合併，重新造一個新的 data 給 `sheet_from_array_of_arrays`

```javascript=
  var data = [];
  for (let i = 0; i < jsonData.length; i++) {
    let arr = [];
    arr.push(th[i]);
    for (let j = 0; j < jsonData[i].length; j++) {
      arr.push(jsonData[i][j])
    }
    data.push(arr);
  }
```

最後在把其他相應的程式碼轉變為迴圈的方式，以及新增傳入分頁名稱的陣列就 OK 了

```javascript=
export function export_json_to_excel(th, jsonData, defaultTitle, sheetArray) {
  // 產生多個分頁
  // 重新組範例的不適合
  var data = [];
  for (let i = 0; i < jsonData.length; i++) {
    let arr = [];
    arr.push(th[i]);
    for (let j = 0; j < jsonData[i].length; j++) {
      arr.push(jsonData[i][j])
    }
    console.log('arr', arr);
    data.push(arr);
  }

  // 定義分頁名稱
  var ws_name = sheetArray;

  var wb = new Workbook(), ws = [];
  // 數據轉換
  for (var j = 0; j < th.length; j++) {
    ws.push(sheet_from_array_of_arrays(data[j]))
  }

  // 產生分頁
  for (var k = 0; k < th.length; k++) {
    wb.SheetNames.push(ws_name[k])
    wb.Sheets[ws_name[k]] = ws[k]
  }

  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: false,
    autoWidth: true,
    type: 'binary'
  });
  var title = defaultTitle || '列表'

  saveAs(new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }), title + ".xlsx")
};
```

## 改變樣式

因為原生的 `js-xlsx` 免費版沒有提供改變 excel 單元格樣式的方法 ( 要付費 )，所以就有人基於此產生了 [xlsx-style](https://www.npmjs.com/package/xlsx-style) 套件，可以改變簡易的樣式，缺點是最近一次的更新已經是四年前了。

### 安裝

```
npm i xlsx-style --save
```
在 `Export2Excel.js` 開頭引入
```javascript=
require('script-loader!file-saver');
require('./Blob.js');
require('script-loader!xlsx/dist/xlsx.core.min');
// 引入樣式
import XLSX from 'xlsx-style'
```

### 設定樣式
官方的首頁內容並沒有手把手教該怎樣設定樣式，只有[樣式設定的表格](https://www.npmjs.com/package/xlsx-style#cell-styles) ~~真的很坑~~
從官方 GitLab [test 資料夾](https://github.com/protobi/js-xlsx/blob/beta/tests/test-style.js#L100)上，可以看出一些端倪，我們應該要如何設定樣式
![](https://trello-attachments.s3.amazonaws.com/5f20e6dc8ba6238874c056e2/5f6ac457c0ce208b7f4d8a0f/770b52a3c831a3960ec5615849bc1258/%E8%A8%BB%E8%A7%A3_2020-10-07_175449.jpg)

這個結構有點似曾相似......

就是先前 [產生分頁 Export2Excel.js](#Export2Exceljs) 圖片的結構，差在物件多了名為 `s` 的屬性，看到關鍵字 `border` ，聯想到剛剛所說的[樣式設定的表格](https://www.npmjs.com/package/xlsx-style#cell-styles)

### 前端控制

接下來就產生了一個問題，樣式該如何插入？

幸好 `Export2Excel.js` 網路上流傳的版本都不是壓縮後的 JS 檔，我們可以從 `function sheet_from_array_of_arrays` 中找到一段改變我們所傳進來的資料結構程式碼

```javascript=
// 網路上 Export2Excel.js 有很多，不確定每個是否一樣
// 大概是在被第二個 for 迴圈包起來一開始的幾行程式碼中
var cell = { v: data[R][C] };
```

我們的任務就是把 `{ v: ... }` 結構轉換的過程也由我們轉換，並且加入 `s` 屬性
```json=
{ v: ...,
  s: { border: { ... } }
}
```

`tableData` 的結構就會相對來說變得有點複雜

```
tableData :[
    [
      {
        id: { v: "四五", s: { border: { top: { style: "thin" } } } },
        title: { v: "一二三", s: { border: { top: { style: "thin" } } } },
        author: { v: 3, s: { border: { top: { style: "thin" } } } },
      },
      {
        id: { v: "四五", s: { border: { top: { style: "thin" } } } },
        title: { v: "一二三", s: { border: { top: { style: "thin" } } } },
        author: { v: 4, s: { border: { top: { style: "thin" } } } },
      },
    ],
    ...
    ...
    ...
]
```

你當然也可以直接寫在 `Export2Excel.js` 裡面，但我本身因為專案的需求必須判斷哪欄只要 `border-top` 、哪個只要 `border-left` ......叭啦叭啦，到最後再插入樣式已經來不及了。

### 輸出失敗

好不容易把資料結構改變好要輸出成 excel 時卻失敗，原因是 `/node_modules/xlsx-style/dist/cpexcel.js` 的第 807 行
```
把 var cpt = require('./cpt' + 'able');
改成 var cpt = cptable;
```
這樣就可以了，很多人遇到這個問題，但作者堅持不改，印象中有搜尋到改的話這個套件會變得相對大(不是很懂)，總之每次 `npm install` 就會遇到這樣的問題，請記住。


## 結語

說真的，有點不太推薦使用這樣的方法來讓 excel 產生樣式，~~對新手來說有點複雜 XD，~~ 如同上面所說的最近一次的更新已經是五年前了，但是就純前端的領域來說目前還沒有找到其他免費方法產生樣式，。

必須靠後端像是 `Node.js` 的 [ExcelJS](https://github.com/exceljs/exceljs) 等套件來完成。

## 範例檔
- [GitHub](https://github.com/Nighthree/excel_worksheets_style)
- [Blob.js](https://github.com/Nighthree/excel_worksheets_style/blob/main/Blob.js)
- [Export2Excel.js](https://github.com/Nighthree/excel_worksheets_style/blob/main/Export2Excel.js)

## 參考
- [Vue实现导出excel](https://juejin.im/post/6844904146076712974)
- [vue导出Excel格式 vue-json-excel file-saver xlsx](https://zhuanlan.zhihu.com/p/66069444)
- [在VUE中使用Export2Excel导出表格，导出多个sheet](https://blog.csdn.net/qq_42651390/article/details/105584685)
- [js-xlsx 单元格边框设置](https://www.hxstrive.com/article/751.htm)
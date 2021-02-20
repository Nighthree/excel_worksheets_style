<template>
  <div>
    <button @click="exportExcel">导出excel</button>
  </div>
</template>
<script>
export default {
  data() {
    return {
      data: [],
      tableDataTest: [
        [
          {
            id: "四五六",
            title: "一二三四",
            author: 3,
            pageviews: 4,
            display_time: 5,
          },
          { id: 6, title: 7, author: 8, pageviews: 9, display_time: 10 },
          { id: 11, title: 12, author: 13, pageviews: 14, display_time: 15 },
        ],
        [
          {
            id: "四五",
            title: "一二三",
            author: 3,
            pageviews: 4,
            display_time: 5,
          },
          { id: 6, title: 7, author: 8, pageviews: 9, display_time: 10 },
          { id: 11, title: 12, author: 13, pageviews: 14, display_time: 15 },
        ],

        [
          {
            id: "四五",
            title: "一二三",
            author: 3,
            pageviews: 4,
            display_time: 5,
          },
          { id: 6, title: 7, author: 8, pageviews: 9, display_time: 10 },
          { id: 11, title: 12, author: 13, pageviews: 14, display_time: 15 },
        ],
      ],
      tableData: [
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
        [
          {
            id: {
              v: "四五",
              s: {
                border: {
                  top: { style: "thin" },
                },
              },
            },
            title: {
              v: "一二三",
              s: {
                border: {
                  top: { style: "thin" },
                },
              },
            },
            author: {
              v: 5,
              s: {
                border: {
                  top: { style: "thin" },
                },
              },
            },
          },
        ],
      ],
      tableDataS: [
        {
          id: {
            v: "四五",
            s: {
              border: {
                top: { style: "thin" },
              },
            },
          },
          title: {
            v: "一二三",
            s: {
              border: {
                top: { style: "thin" },
              },
            },
          },
          author: {
            v: 4,
            s: {
              border: {
                top: { style: "thin" },
              },
            },
          },
        },
      ],
    };
  },
  methods: {
    //导出的方法
    exportExcel() {
      require.ensure([], () => {
        // 產生多個分頁
        const { export_json_to_excel } = require("../Excel/Export2Excel"); //注意这个Export2Excel路径
        const tHeader = [
          ["序號", "暱稱", "姓名"],
          ["序號", "暱稱", "姓名"],
          ["序號", "暱稱", "姓名"],
        ]; // 上面设置Excel的表格第一行的标题 , { v: "昵称" }, { v: "姓名" } [{ v: "序号1" }, { v: "序号2" }];
        const filterVal = [
          ["id", "title", "author"],
          ["id", "title", "author"],
          ["id", "title", "author"],
        ]; // 上面的index、nickName、name是tableData里对象的属性key值
        const list = this.tableDataTest; //把要导出的数据tableData存到list
        // console.log(list);
        // const data = this.formatJson(filterVal, list);
        let data = [];
        for (let i = 0; i < list.length; i++) {
          let arr = this.formatJson(filterVal[i], list[i]);
          data.push(arr);
        }

        // 分頁名稱
        const sheetArray = ["Sheet1", "Sheet2", "Sheet3"];
        // 產生一個分頁
        // const { export_json_to_excel } = require("../Excel/Export2Excel"); //注意这个Export2Excel路径
        // const tHeader = [{ v: "序号" }]; // 上面设置Excel的表格第一行的标题 , { v: "昵称" }, { v: "姓名" }
        // const filterVal = ["id", "title", "author"]; // 上面的index、nickName、name是tableData里对象的属性key值
        // const list = this.tableDataS; //把要导出的数据tableData存到list
        // const data = this.formatJson(filterVal, list);

        export_json_to_excel(tHeader, data, "导出文件名", sheetArray); //最后一个是表名字
      });
    },
    formatJson(filterVal, jsonData) {
      return jsonData.map((v) => filterVal.map((j) => v[j]));
    },
  },
};
</script>
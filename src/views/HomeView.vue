
<template>
  <div class="home">
    <button @click="exportExcel">导出execl</button>
    <div class="tableContainer">
      <table id="tableContainer">
        <thead>
          <tr>
            <td colspan="8">机房操作表单</td>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>设备编号</td>
            <td colspan="7">123456</td>
          </tr>
        </tbody>
        <tbody>
          <tr>
            <td>设备描述</td>
            <td colspan="7">用于冷却计算机用于冷却计算机用于冷却计算机用于冷却计算机用于冷却计算机用于冷却计算机用于冷却计算机用于冷却计算机用于冷却计算机</td>
          </tr>
        </tbody>
        <tbody>
          <tr>
            <td colspan="8">设备信息</td>
          </tr>
          <tr>
            <td>设备序号</td>
            <td colspan="7">内容</td>
          </tr>
        </tbody>

      
      </table>
    </div>
  </div>
</template>

<script>
const XLSX = window.XLSX;
const XLSXX = window.XLSXX;
export default {
  name: "HomeView",
  data() {
    return {};
  },
  components: {},
  methods: {
    exportExcel() {
      this.sheetJSExportTable(document.getElementById("tableContainer"));
    },

    sheetJSExportTable(table) {
      var elt = table;
      var wb = XLSX.utils.table_to_book(elt, { sheet: "Sheet JS", raw: true });
      let range = XLSX.utils.decode_range(wb.Sheets["Sheet JS"]["!ref"]);
      console.log(wb);
      console.log(range);
      let borderStyle = {
        top: {
          style: "thin",
          color: { rgb: "000000" },
        },
        bottom: {
          style: "thin",
          color: { rgb: "000000" },
        },
        left: {
          style: "thin",
          color: { rgb: "000000" },
        },
        right: {
          style: "thin",
          color: { rgb: "000000" },
        },
      };
      wb.Sheets["Sheet JS"]["!merges"].forEach((item) => {
        if (item.e.r == item.s.r && item.e.c != item.s.c) {
          // 列合并
          let R = item.s.r;
          // let countLength = item.e.c - item.s.c;
          for (let i = item.s.c; i <= item.e.c; i++) {
            let cell = { c: i, r: R };
            let cell_ref = XLSX.utils.encode_cell(cell);
            if (!wb.Sheets["Sheet JS"][cell_ref]) {
              wb.Sheets["Sheet JS"][cell_ref] = { t: "s", v: "" };
            }
          }
        } else if (item.e.c == item.s.c && item.e.r != item.s.r) {
          // 行合并
          let C = item.s.c;
          // let countLength = item.e.r - item.s.r;
          for (let i = item.s.r; i <= item.e.r; i++) {
            let cell = { c: C, r: i };
            let cell_ref = XLSX.utils.encode_cell(cell);
            if (!wb.Sheets["Sheet JS"][cell_ref]) {
              wb.Sheets["Sheet JS"][cell_ref] = { t: "s", v: "" };
            }
          }
        }
      });
      console.log(wb);
      console.log(range);
      for (let C = range.s.c; C <= range.e.c; ++C) {
        for (let R = range.s.r; R <= range.e.r; ++R) {
          let cell = { c: C, r: R };
          let cell_ref = XLSX.utils.encode_cell(cell);
          if (R == 1) {
            wb.Sheets["Sheet JS"][cell_ref].s = {
              alignment: {
                horizontal: "center",
                vertical: "center",
              },
              font: {
                name: "黑体",
                // sz: "15",
                bold: true,
              },
              border: borderStyle,
            };
          } else {
            if (wb.Sheets["Sheet JS"][cell_ref]) {
              wb.Sheets["Sheet JS"][cell_ref].s = {
                font: {
                  name: "黑体",
                  // sz: "12",
                },
                border: borderStyle,
              };
            }
          }
        }
      }
      var wopts = { bookType: "xlsx", bookSST: false, type: "binary" };
      var wbout = XLSXX.write(wb, wopts); // 使用xlsx-style 写入
      // eslint-disable-next-line
      saveAs(new Blob([this.s2ab(wbout)], { type: "" }), "aa.xlsx");
    },

    s2ab(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
      return buf;
    },
  },
};
</script>

<style scoped>
.tableContainer{
  width: 100vw;
  height: 100vh;
}
table {
  border-collapse: collapse;
  border: 1px solid #ccc;
}
th, td {
  padding: 10px;
  border: 1px solid #ccc;
}
th {
  background-color: #ccc;
}
</style>

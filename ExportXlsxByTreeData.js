/*
 * @Author: bucai
 * @Date: 2021-05-05 22:31:47
 * @LastEditors: bucai
 * @LastEditTime: 2021-05-06 09:32:03
 * @Description: 
 */
const ExportXlsxByTreeData = {
  // 保存新的数据 直接作用于导出  
  newData: {},
  // 合并单元格时用
  merges: [],
  // 内部合并单元格所需数据
  startNObj: {},
  // 默认开始的number 如果包含header 那就是 startNDefault = 2 否则就是1
  startNDefault: 2,
  /**
   * 保存数据
   * @param {*} key 
   * @param {*} value 
   */
  _saveData (key, value) {
    this.newData[key] = {
      v: String(value),
      s: {
        alignment: {
          vertical: "middle",
          horizontal: "center",
          wrapText: true
        },
      }
    };
  },
  /**
   * 保存合并的单元位置
   * @param {*} location 
   * @param {*} len 
   */
  _saveMerge (location, len) {
    this.merges.push({
      s: {
        c: location.c,
        r: location.r,
      },
      e: {
        c: location.c,
        r: location.r + len
      }
    })
  },
  /**
   * 内部函数，无需理会
   * @param {*} data 数据
   * @param {*} start 
   * @returns 
   */
  _parse (data = [], startX = 'A') {
    // 做一个初始化
    this.startNObj[startX] = this.startNObj[startX] || this.startNDefault;
    // 得到当前位置
    let startN = this.startNObj[startX];
    // 子项的高度
    let listLength = 0;
    // 解析数组
    data.forEach(item => {
      // hack 处理
      const list = item.list || [];
      // 当前list长度
      let currentLen = 1;
      if (list.length) {
        // 如果存在子集
        const nl = this._parse(list, String.fromCodePoint(startX.codePointAt(0) + 1));
        // 如果存在子项，那么len就等于
        listLength += nl;
        currentLen = nl;

      } else {
        listLength += 1;
      }
      // 保存单个节点
      const key = startX + startN;
      const value = item.value;
      this._saveData(key, value)

      // 如果只有一个单元格就不需要处理
      if (currentLen > 1) {
        // 计算位置
        // c = col
        // r = row
        const startC = startX.codePointAt(0) - 'A'.codePointAt(0);
        const len = currentLen - 1;
        // 保存合并单元格
        this._saveMerge({ c: startC, r: startN - 1 }, len);
      }
      // 更新位置
      startN = (this.startNObj[startX] += currentLen);
      return item;
    });
    return listLength;
  },
  _init () {
    this.startNDefault = 2;
    this.merges = [];
    this.newData = {};
    this.startNObj = {};
  },

  /**
   * 解析数据
   * @param {Array} header 头信息
   * @param {Array} treeData 树状数据
   * @example 
   *  exportExcel(["用餐日期", "餐别", "地址", "工号", "员工姓名", "菜品名称", "份数"],[
   *    {
   *      value: "显示的值",
   *      list: [
   *        {
   *          value: "显示的值",
   *        }
   *      ]
   *    }
   *  ])
   */
  exportExcel (header, treeData) {
    this._init();

    treeData = JSON.parse(JSON.stringify(treeData))
    this._parse(treeData);
    // 合并 headers 和 data
    // const header = ["用餐日期", "餐别", "地址", "工号", "员工姓名", "菜品名称", "份数"];
    var output = Object.assign({}, header.reduce((pv, cv, i) => {
      pv[String.fromCodePoint('A'.codePointAt() + i) + '1'] = {
        v: cv,
        s: {
          alignment: {
            horizontal: 'center'
          },
          font: {
            bold: true
          }
        }
      };
      return pv;
    }, {}), this.newData);
    // 表格范围，范围越大生成越慢
    // 构建 workbook 对象
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    worksheet.properties.defaultColWidth = 14;
    // 解析数据
    Object.keys(output).forEach(key => {
      const { v, s } = output[key];
      const cell = worksheet.getCell(key);
      cell.value = v;
      if (s) {
        cell.alignment = s.alignment;
        cell.fill = s.fill;
        cell.font = s.font;
      }
    });
    this.merges.forEach(({ s, e }) => {
      const sA = String.fromCodePoint('A'.codePointAt(0) + s.c) + (s.r + 1);
      const sB = String.fromCodePoint('A'.codePointAt(0) + e.c) + (e.r + 1);
      worksheet.mergeCells(sA + ':' + sB);
    });

    workbook.xlsx.writeBuffer().then(function (data) {
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8' });
      saveAs(blob, 'tests.xlsx');
    });
  }
}

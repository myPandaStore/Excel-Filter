function fileToJson(file, callback) {
  // 数据处理结果
  let result;
  // 是否用BinaryString（字节字符串格式） 否则使用base64（二进制格式）
  let isBinary = true;

  // 读取本地 Excel 文件
  var reader = new FileReader();
  // debugger
  reader.onload = function (e) {
    var data = e.target.result;
    // debugger
    if (isBinary) {
      result = XLSX.read(data, {
        type: "binary",
        cellDates: true,
      });
      // debugger
    } else {
      result = XLSX.read(btoa(fixdata(data)), {
        type: "base64",
        cellDates: true,
      });
    }
    // 格式化数据
    formatResult(result, callback);
  };
  if (isBinary) {
    reader.readAsBinaryString(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

function formatResult(data, callback) {
  // 获取总数据
  const sheets = data.Sheets;
  // debugger
  // 获取每个表格
  const sheetItem = Object.keys(sheets);
  // debugger
  // 返回sheetJSON数据源
  let sheetArr = [];
  // 获取
  sheetItem.forEach((item) => {
    const sheetJson = XLSX.utils.sheet_to_json(sheets[item], { header: 1 });
    // debugger
    // 格式化Item时间数据
    formatItemDate(sheetJson);
    // 将名词复数转变成单数并去重（补充去重逻辑）
    NounPluralizeToSingularize(sheetJson);
    // debugger;
    // 格式化Item合并数据
    formatItemMerge(sheets[item], sheetJson);
    // 组合数据
    sheetArr.push({
      name: item,
      list: sheetJson,
    });
  });
  // 返回数据
  callback(sheetArr);
}

// list 存储最后整理完毕需要导出的数据
let list = [];
function NounPluralizeToSingularize(sheetJson) {
  let newArr = [], size = 10
  for (let i = 0; i < sheetJson.length; i++) {
    // 当前行
    let currentRow = sheetJson[i];
    for (let j = 0; j < currentRow.length; j++) {
      if (typeof currentRow[j] === 'number') {
        currentRow[j] = String(currentRow[j])
      }
      let phraseArr = currentRow[j].split(" ");
      for (let k = 0; k < phraseArr.length; k++) {
        if (phraseArr[k] == "plus" || phraseArr[k] == "towards") {
          phraseArr[k] = phraseArr[k];
        } else {
          phraseArr[k] = Inflector.singularize(phraseArr[k]);
        }
      }
      let res = phraseArr.join(" ");
      newArr.push(res);
    }
  }
  newArr = unique(newArr);
  newArr = upSort(newArr);
  size = Number(window.localStorage.getItem('size')) || size
  newArr = sliceArray(newArr, size)
  for (let i = 0; i < newArr.length; i++) {
    let newObj = Object.assign({}, newArr[i]);
    list.push(newObj)
  }
}
function chooseLength(e) {
  let value = e.target.value
  window.localStorage.setItem("size", value)
}
// 数组拆分
function sliceArray(arr,size) {
  var newArr = []
  for (var i = 0; i < arr.length; i=i+size) {
    newArr.push(arr.slice(i,i+size))
  }
  return newArr
}

// 去重
function unique(arr) {
  return Array.from(new Set(arr));
}
// 升序排序
function upSort(arr) {
  return arr.sort((a, b) => a.length - b.length);
}
function formatItemDate(data) {
  data.forEach((row) => {
    row.forEach((item, index) => {
      // 若有数据为时间格式则格式化时间
      if (item instanceof Date) {
        // 坑：这里因为XLSX插件源码中获取的时间少了近43秒，所以在获取凌晨的时间上会相差一天的情况,这里手动将时间加上
        var date = new Date(Date.parse(item) + 43 * 1000);
        row[index] = `${date.getFullYear()}-${String(
          date.getMonth() + 1
        ).padStart(2, 0)}-${String(date.getDate()).padStart(2, 0)}`;
      }
    });
  });
}

function formatItemMerge(sheetItem, data) {
  const merges = sheetItem["!merges"] || [];
  merges.forEach((el) => {
    const start = el.s;
    const end = el.e;
    // 处理行合并数据
    if (start.r === end.r) {
      const item = data[start.r][start.c];
      for (let index = start.c; index <= end.c; index++) {
        data[start.r][index] = item;
      }
    }
    // 处理列合并数据
    if (start.c === end.c) {
      const item = data[start.r][start.c];
      for (let index = start.r; index <= end.r; index++) {
        data[index][start.c] = item;
      }
    }
  });
}

// 文件流转 base64
function fixdata(data) {
  var o = "",
    l = 0,
    w = 10240;
  for (; l < data.byteLength / w; ++l)
    o += String.fromCharCode.apply(
      null,
      new Uint8Array(data.slice(l * w, l * w + w))
    );
  o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
  return o;
}


// 导出数据
function ExportData() {
  //相关连接http://www.qinluo1023.com/2018/07/03/JS%E8%AF%BB%E5%8F%96%E6%9C%AC%E5%9C%B0EXCEL%E6%96%87%E4%BB%B6.html
  var data = list;

  /* 创建worksheet */
  var ws = XLSX.utils.json_to_sheet(data, {
    skipHeader: true,
  });

  /* 新建空workbook，然后加入worksheet */
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "sheetjs");

  /* 生成xlsx文件 */
  XLSX.writeFile(wb, "sheetjs.xlsx");
}

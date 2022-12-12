function fileToJson(file, callback) {
  // 数据处理结果
  let result;
  // 是否用BinaryString（字节字符串格式） 否则使用base64（二进制格式）
  let isBinary = true;

  // 读取本地 Excel 文件
  var reader = new FileReader();
  reader.onload = function (e) {
    var data = e.target.result;
    if (isBinary) {
      result = XLSX.read(data, {
        type: "binary",
        cellDates: true,
      });
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
  // 获取每个表格
  const sheetItem = Object.keys(sheets);
  // 返回sheetJSON数据源
  let sheetArr = [];
  // 获取
  sheetItem.forEach((item) => {
    const sheetJson = XLSX.utils.sheet_to_json(sheets[item], {
      header: 1,
      defval: ''
    });
    let transposeSheetJson = transpose(sheetJson)
    // 格式化Item时间数据
    formatItemDate(sheetJson);
    // 将名词复数转变成单数并去重（补充去重逻辑）
    NounPluralizeToSingularize(transposeSheetJson);
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

// finalData 存储最后整理完毕需要导出的数据
let finalData = []
function NounPluralizeToSingularize(sheetJson) {
  // 获取最长列的数据长度
  let sheetJsonLength = sheetJson[0].length
 
  // 遍历每一行
  for (let i = 0; i < sheetJson.length; i++) {
    let newArr = []
    let currentRow = sheetJson[i];
    let phraseArr = []
    let res
    // 遍历每一个短语
    for (let j = 0; j < currentRow.length; j++) {
      if (typeof currentRow[j] === 'number') {
        currentRow[j] = String(currentRow[j])
      }
      // 将当前行的当前 phrase 按照空格进行分割
      phraseArr = currentRow[j].split(" ");
      // 遍历当前 phrase 的每一个单词
      for (let k = 0; k < phraseArr.length; k++) {
        // 特殊介词不处理
        if (phraseArr[k] == "plus" || phraseArr[k] == "towards" || phraseArr[k] == 'is') {
          phraseArr[k] = phraseArr[k];
        } else {
          // 名词复数转单数
          phraseArr[k] = Inflector.singularize(phraseArr[k]);
        }
      }
      res = phraseArr.join(" ");
      newArr.push(res)
    }
    // 当前 excel 列处理好的数据进行去重排序处理（一列只存储特定品类数据）
    newArr = upSort(newArr);
    newArr = unique(newArr).slice(1);
    // 添加空串占位
    let addEmptyStringArr = new Array(sheetJsonLength - newArr.length).fill('')
    newArr = newArr.concat(addEmptyStringArr)
    finalData.push(newArr)
  }
  finalData = transpose(finalData)
  for (let i = 0; i < finalData.length; i++) {
    finalData[i] = Object.assign({},finalData[i])
  }
}

// 补充空串
function replenishEmptyString(newArr,sheetJsonLength) {
  let len = newArr.length
  for (let i = len; len < sheetJsonLength; i++) {
    newArr.push('')
  }
  return newArr
}
// 二维数据行列转换
function transpose(arr) {
  var newArray = arr[0].map(function (col, i) {
    return arr.map(function (row) {
      return row[i];
    })
  });
  return newArray
}

// 数组拆分
function sliceArray(arr, size) {
  var newArr = []
  for (var i = 0; i < arr.length; i = i + size) {
    newArr.push(arr.slice(i, i + size))
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
  var data = finalData;

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
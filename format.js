function fileToJson(file, callback) {
 // 数据处理结果
 let result;
 // 是否用BinaryString（字节字符串格式） 否则使用base64（二进制格式）
 let isBinary = true;

 var reader = new FileReader();
 reader.onload = function (e) {
  var data = e.target.result;
  if (isBinary) {
   result = XLSX.read(data, {
    type: 'binary',
    cellDates: true,
   });
  } else {
   result = XLSX.read(btoa(fixdata(data)), {
    type: 'base64',
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
  const sheetJson = XLSX.utils.sheet_to_json(sheets[item], { header: 1 });
  // 格式化Item时间数据
  formatItemDate(sheetJson);
  // 格式化Item合并数据
  formatItemMerge(sheets[item], sheetJson);
  // 组合数据
  sheetArr.push({
   name: item,
   list: sheetJson
  })
 });
 // 返回数据
 callback(sheetArr)
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
 // console.log('data', data);
}

function formatItemMerge(sheetItem, data) {
 const merges = sheetItem['!merges'] || [];
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
 var o = '',
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
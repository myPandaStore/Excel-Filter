const file = document.querySelector('input');
function changes(e) {
  if (e.target.files.length > 0) {
    const fileName = e.target.files[0].name;
    const fileArr = fileName.split('.');
    const fileSuffix = fileArr[fileArr.length - 1];
    if (fileSuffix === 'xlsx' || fileSuffix === 'xls') {
      fileToJson(e.target.files[0], (sheets) => {
        console.log('获取到的表格数据', sheets);
      });
    } else {
      console.log('不支持该格式的解析');
    }
  } else {
    console.log('请选择文件上传');
  }
}
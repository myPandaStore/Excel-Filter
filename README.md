# XLSX
XLSX获取Excel内容
工作中经常遇到让前端来识别Excel的内容，以此来进行校验，那么首先如何获取Excel的内容就是一个问题，百度了一下发现用SheetJS开源的JS-XLSX这个工具库居多。它支持xls、xlsx、csv等多种表格格式文件的解析。
demo图例：https://cdn.nlark.com/yuque/0/2021/png/1485818/1609897596492-0ce714e3-bf3f-43aa-b8e2-89f9b45681ae.png
最后期望获取的数据结构：https://cdn.nlark.com/yuque/0/2021/png/1485818/1609897660281-c534795d-a118-489f-bdb0-7798000c119e.png
官方github：https://github.com/SheetJS/js-xlsx

• 如何使用XLSX
首先，官网这里的文件只需要用dist目录下的 xlsx.core.min.js 就足够了；xlsx.core.min.js主要是包含了基本的识别解析功能，xlsx.full.min.js则是包含了所有功能模块。然后在所在项目引入即可。
下面演示的案例是在原生JS下写的，所以用的script标签引入的，若是Vue等框架项目可以用 npm 下载引入 XLSX即可。

![image](https://cdn.nlark.com/yuque/0/2021/png/1485818/1609897596492-0ce714e3-bf3f-43aa-b8e2-89f9b45681ae.png)

# EXCEL FILTER
Excel Filter 使用说明：
  1.下载 vscode代码编辑，导入源代码，在编辑器内鼠标右键点击打开本地浏览器
  2.输入想要导出的每一行 excel 数据长度(默认为10，可以不输入,请输入有效的数字)
  注意：每次要选择不同的长度时，点击网页刷新缓存，在导入导出文件
  3.点击选择文件按钮（导入本地的excel文件，仅支持 xlsx 或 xls文件）。
  注意：若想要保证每次导入导出干净的新的excel文件，记得点击左上角刷新缓存。否则生成
  的文件会包含历史缓存文件。如视频示范所示。
  4.点击导出数据会自动生成过滤后的 excel 表格（按照 phrase 长度升序排列）
  5.暂时不支持单数转复数，会有更多的边界条件需要处理，开发量大。
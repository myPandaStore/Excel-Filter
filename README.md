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

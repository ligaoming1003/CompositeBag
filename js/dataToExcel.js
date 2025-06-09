// Javascrīpt Document
  // 杀Excel进程使用此变量
  var idTmr = "";
  // 函数功能：复制表格到Excel中
  // 参    数：tableID  表的id
  function CellToTable(tableID)
  {
        try
        {
                  // 加载ActiveX控件，获取Excel句柄
                  var exApp = new ActiveXObject("Excel.Application");
                // 创建一个Excel文件
                var ōWB = exApp.WorkBooks.add();
                // 获取sheet1句柄CA
                var exSheet = exApp.ActiveWorkBook.WorkSheets(1);
// 设置sheet1的名称
exSheet.name="演示复制表格到Excel中";
                // copy指定的表格
                var sel=document.body.createTextRange();
            sel.moveToElementText(tableID);
            sel.select();
            sel.execCommand("Copy");
                // 粘贴到sheet中
                exSheet.Paste();
                // 弹出保存对话框，保存Excel文件 
                exApp.Save();
                // 退出Excel实例
                exApp.Quit();
                exApp = null;
                // 调用Cleanup（）进行垃圾回收
                idTmr = window.setInterval("Cleanup();",10);
        }catch(Exception)
        {
                //用户点击保存对话框中的取消按钮时会发生异常
        }
  }
  // 函数功能：杀掉Excel进程
  function Cleanup() {
          window.clearInterval(idTmr);
          CollectGarbage();
  }
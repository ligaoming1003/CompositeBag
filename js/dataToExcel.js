// Javascr��pt Document
  // ɱExcel����ʹ�ô˱���
  var idTmr = "";
  // �������ܣ����Ʊ��Excel��
  // ��    ����tableID  ���id
  function CellToTable(tableID)
  {
        try
        {
                  // ����ActiveX�ؼ�����ȡExcel���
                  var exApp = new ActiveXObject("Excel.Application");
                // ����һ��Excel�ļ�
                var ��WB = exApp.WorkBooks.add();
                // ��ȡsheet1���CA
                var exSheet = exApp.ActiveWorkBook.WorkSheets(1);
// ����sheet1������
exSheet.name="��ʾ���Ʊ��Excel��";
                // copyָ���ı��
                var sel=document.body.createTextRange();
            sel.moveToElementText(tableID);
            sel.select();
            sel.execCommand("Copy");
                // ճ����sheet��
                exSheet.Paste();
                // ��������Ի��򣬱���Excel�ļ� 
                exApp.Save();
                // �˳�Excelʵ��
                exApp.Quit();
                exApp = null;
                // ����Cleanup����������������
                idTmr = window.setInterval("Cleanup();",10);
        }catch(Exception)
        {
                //�û��������Ի����е�ȡ����ťʱ�ᷢ���쳣
        }
  }
  // �������ܣ�ɱ��Excel����
  function Cleanup() {
          window.clearInterval(idTmr);
          CollectGarbage();
  }
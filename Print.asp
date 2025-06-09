<script defer>  
　　function SetPrintSettings() {  
　　// -- advanced features  
　　factory.printing.SetMarginMeasure(2) // measure margins in inches  
　　factory.SetPageRange(false, 1, 3) // need pages from 1 to 3  
　　factory.printing.printer = "HP DeskJet 870C"  
　　factory.printing.copies = 2  
　　factory.printing.collate = true  
　　factory.printing.paperSize = "A4"  
　　factory.printing.paperSource = "Manual feed"  

　　// -- basic features  
　　factory.printing.header = "This is MeadCo"  
　　factory.printing.footer = "Advanced Printing by ScriptX"  
　　factory.printing.portrait = false  
　　factory.printing.leftMargin = 1.0  
　　factory.printing.topMargin = 1.0  
　　factory.printing.rightMargin = 1.0  
　　factory.printing.bottomMargin = 1.0  
　　}  
</script>

<script language="javascript">  
　　function printsetup(){  
　　// 嬉咫匈中譜崔  
　　wb.execwb(8,1);  
　　}  
　　function printpreview(){  
　　// 嬉咫匈中圓誓  
　　
　　wb.execwb(7,1);  
　　　
　　}  

　　function printit()  
　　{  
　　if (confirm('鳩協嬉咫宅��')) {  
　　wb.execwb(6,6)  
　　}  
　　}  
</script>
<OBJECT classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" height=0 id=wb name=wb width=0></OBJECT>  
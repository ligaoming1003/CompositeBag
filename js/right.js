var menuload="<table id=rightmenu width='85' border='1' cellspacing='0' style='padding-left:16;padding-top:1;padding-bottom:1; width:84px; z-index:50;visibility: hidden;' class='window-0'>"
+"<tr><td class='window-1'><table width='100%' border='0' cellpadding='3' cellspacing='0'>"
+"<tr><td align='center' onClick=document.execCommand('Undo')  onMouseOver=this.style.background='#08246B';this.style.color='#FFFFFF' onMouseOut=this.style.background='';this.style.color=''>³·Ïû(<U>Z</U>)</td>"
+"</tr><tr><td align='center'  onClick=document.execCommand('Cut') onMouseOver=this.style.background='#08246B';this.style.color='#FFFFFF' onMouseOut=this.style.background='';this.style.color=''>¼ôÇÐ(<U>T</U>)</td>"
+"</tr><tr><td align='center' onClick=document.execCommand('Copy')  onMouseOver=this.style.background='#08246B';this.style.color='#FFFFFF' onMouseOut=this.style.background='';this.style.color=''>¸´ÖÆ(<U>C</U>)</td>"
+"</tr><tr><td align='center' onClick=document.execCommand('Paste')  onMouseOver=this.style.background='#08246B';this.style.color='#FFFFFF' onMouseOut=this.style.background='';this.style.color=''>Õ³Ìù(<U>P</U>)</td>"
+"</tr><tr><td align='center' onClick=document.execCommand('Delete')  onMouseOver=this.style.background='#08246B';this.style.color='#FFFFFF' onMouseOut=this.style.background='';this.style.color=''>É¾³ý(<U>D</U>)</td>"
+"</tr><tr><td align='center' onClick=this.select()   onMouseOver=this.style.background='#08246B';this.style.color='#FFFFFF' onMouseOut=this.style.background='';this.style.color=''>È«Ñ¡(<u>A</u>)</td>"
+"</tr></table></td></tr></table>"
document.write(menuload);
function showmenuie5(){
 var rightedge=document.body.clientWidth-event.clientX-100
 var bottomedge=document.body.clientHeight-event.clientY-25
if (rightedge<rightmenu.offsetWidth)
 rightmenu.style.left=document.body.scrollLeft+event.clientX-rightmenu.offsetWidth;
else
 rightmenu.style.left=document.body.scrollLeft+event.clientX
if (bottomedgerightmenu.offsetHeight)
 rightmenu.style.top=document.body.scrollTop+event.clientY-rightmenu.offsetHeight
else
 rightmenu.style.top=document.body.scrollTop+event.clientY
 rightmenu.style.visibility="visible"
return false
}
function hidemenuie5(){
 rightmenu.style.visibility="hidden"
 }
document.oncontextmenu=showmenuie5;
document.body.onclick=hidemenuie5;

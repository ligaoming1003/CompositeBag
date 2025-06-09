function document.oncontextmenu(){
if ((event.srcElement.type=="text")||(event.srcElement.type=="textarea"))
{return true;}else{return false;}
}
function document.ondragstart(){
if ((event.srcElement.type=="text")||(event.srcElement.type=="textarea"))
{return true;}else{return false;}
}
function document.onselectstart(){
if ((event.srcElement.type=="text")||(event.srcElement.type=="textarea")||(event.srcElement.className=="textarea"))
{return true;}else{return false;}
}
function document.onmouseover(){
if ((event.srcElement.type=="button")||(event.srcElement.type=="submit")||(event.srcElement.type=="reset")){
if(event.srcElement.className=="button")
{event.srcElement.className="button-o";}else{event.srcElement.className="button1-o";}
}
}
function document.onmouseout(){
if ((event.srcElement.type=="button")||(event.srcElement.type=="submit")||(event.srcElement.type=="reset")){
if(event.srcElement.className=="button-o")
{event.srcElement.className="button";}else{event.srcElement.className="button1";}
}
}
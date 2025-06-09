function defaul_homepage(){
this.homepage.style.behavior='url(#default#homepage)';this.homepage.setHomePage('http://www.qingzhou.net.cn/');
}
function shoucang(){
window.external.AddFavorite('http://www.qingzhou.net.cn/','北京轻舟世纪建筑装饰工程有限公司');
}
function MM_showHideLayers() {
  var i, visStr, args, theObj;
  args = MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) {
  visStr= args[i+2];
   if (navigator.appName == 'Netscape' && document.layers != null) {
      theObj = eval(args[i]);
      if (theObj) theObj.visibility = visStr;
    } else if (document.all != null) { 
   if (visStr == 'show') visStr = 'visible'; 
   if (visStr == 'hide') visStr = 'hidden';
     theObj = eval(args[i+1]);
      if (theObj) theObj.style.visibility = visStr;
 } }
}
function KBAddFavorite(theTit) {
IE4=(document.all)? 1 : 0
if (IE4) {
theSite=window.location
window.external.AddFavorite(theSite,theTit)
}
}
function MM_openBrWindow(theURL,winName,features) {
  window.open(theURL,winName,features);
}
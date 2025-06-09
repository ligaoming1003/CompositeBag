function showtime(){
  var now=new Date()
  var hours=now.getHours()
  var minutes=now.getMinutes()
  var seconds=now.getSeconds()
  var timeValue=""+((hours>12)?hours-12:hours)
  timeValue +=((minutes<10)?":0":":")+minutes
  timeValue +=((seconds<10)?":0":":")+seconds
  timeValue +=(hours>=12)?"下午":"上午"
  //document.kemoon.face2.value="现在时间"+timeValue
  document.all("thename").innerText=timeValue
  setTimeout("showtime()",1000)
}  


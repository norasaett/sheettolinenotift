function myFunction() {
var SheetID = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";   // Google Sheet ID
  var dataSheetID = SpreadsheetApp.openById(SheetID);
  var sheet = dataSheetID.getSheetByName("Sheet1");
  var startRow = 2;
  var numRows = sheet.getDataRange().getLastRow();
  var lastCols = sheet.getDataRange().getLastColumn();
  var dataRange = sheet.getRange(startRow,1, numRows - 1, lastCols);
  var data = dataRange.getValues();
  
  function toborderletter(row,x){          // Function for number css style
     var sp ='';
        for(let i=0;i<=row[x].toString().length-1; i++)
        {
            //Logger.log(row[1].toString()[i]);
            sp+="<span  class='pwn'>" + row[x].toString()[i];
            sp+="</span>&nbsp;";
        }
    return sp;
  }

  for (let i = 0; i < data.length; ++i) {
    var row = data[i];
    var sent = row[5];
    
    //Create html data when not appear sent word in last row
    if (sent != "sent"&&((startRow + i)==numRows)) {
         
 //Use http://pojo.sodhanalibrary.com/string.html to Convert html to javascript variable
 var myvar = '<style>'+
'#p1{'+
'  width: 50px;'+
'  border-radius: 50%;'+
'}'+
'#p2{'+
'  position: absolute;'+
'  width: 50px;'+
'  top: -15px;'+
'  border-radius: 50%;'+
'  margin-left: 160px;'+
'}'+
'.headertitleen{'+
'  position: absolute;'+
'  top: 100%;'+
'  left: 54%;  '+
'  transform: translateX( -50% );'+
'}'+
'.headertitle{'+
'  position: relative;'+
'  top: -10px;'+
'  @border: 15px solid green;'+
'  margin-left: 40px;'+
'  padding-left: 130px;'+
'  padding-right: 100px;'+
'}'+
'table, th, td {'+
'  '+
'   border-collapse: collapse;'+
'}'+
'tr th{'+
'  height: 60px;'+
'  border: 2px solid green;'+
'}'+
'th, td {'+
'  padding: 15px;'+
'  text-align: left;'+
'}'+
''+
'.data{'+
'  border: 3px solid #73AD21;'+
'}'+
''+
'table{'+
'  border: 2px solid green;'+
'  '+
'}'+
''+
'.pwn{'+
' font-size:18px;'+
' font-weight:bold;'+
' padding-left:0px;'+
' padding:2px;'+
' border: 2px solid black ;'+
' display: inline;  '+
'}'+
'.pwn2{'+
' font-size:18px;'+
' font-weight:bold;'+
' padding:2px;'+
' border: 2px solid black ;'+
' display: inline;  '+
'}'+
'</style>'+
'<table>'+
'  <tr>'+
'    <th colspan="3"><img id="p1" src="https://drive.google.com/uc?export=view&id=1b-n0rMY2SIg9-4zLBxVY6G6qyhcSDRl-" /><span class="headertitle">บันทึกความปลอดภัย<span class="headertitleen">SAFTY RECORD</span><img id="p2" src="https://drive.google.com/uc?export=view&id=1J0Vy-cdP3hXXGjH2nQvutGV5d2QsiqEG" /></th>'+
'  </tr>'+
''+
'  <tr>'+
'    <div class="target">'+
'      <td> <span class="title">เป้าหมาย<div>TARGET</div></span></td>'+
'      <td> <span class="number">'+toborderletter(row,1)+'</span></td>'+
''+
'      <td> <span class="unit">ชั่วโมงการทำงาน<div>(MANHOUR)</div></span></td>'+
''+
'    </div>'+
''+
'  </tr>'+
'  <tr>'+
''+
'    <div class="statnum">'+
'      <td> <span class="title">สถิติที่ดีที่สุดในอดีต<div>PAST BEST RECORD</div></span></td>'+
'      <td> <span class="number2">'+toborderletter(row,2)+'</span></td>'+
''+
'      <td> <span class="unit">ชั่วโมงการทำงาน<div>(MANHOUR)</div></span></td>'+
''+
'    </div>'+
'  </tr>'+
'  <tr>'+
'    <div class="target">'+
'      <td> <span class="title">สถิติปัจจุบัน<div>CURRENCY RECORD</div></span></td>'+
'      <td> <span class="number3">'+toborderletter(row,3)+'</span></td>'+
''+
'      <td> <span class="unit">ชั่วโมงการทำงาน<div>(MANHOUR)</div></span></td>'+
''+
'    </div>'+
'  </tr>'+
'  <tr>'+
'    <div class="target">'+
'      <td> <span class="title">อุบัติเหตุครั้งสุดท้ายเมื่อ<div>LAST ACCIDENTAL OCCURENCE</div></span></td>'+
'      <td> <span class="number4">'+row[4]+'</span></td>'+
'      <td> <span class="unit">ข้อมูล ณ '+row[0]+'<div>REPORT DATE</div></span></td>'+
'    </div>'+
'  </tr>'+
'  '+
'</table>';
      
 
      //Record "Sent" in google sheet 
      sheet.getRange(startRow + i, lastCols).setValue("sent");
    }
  }


//Prepre Data before sent to htmltoimage api
//Read api doc from https://docs.htmlcsstoimage.com/ for more info.
var formData = {
  html: myvar
  //You can get screenshot website by sent your desire url 
  //url:'https://script.google.com/macros/s/AKfycbwD4GasFDS9SHqLKHKsXX1896khWRpGNykStRz1TsqvWCmXkYGNXfXuazlzIRrohGqj/exec',
  
};

username = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';  //API KEY FROM htmlcsstoimage.com free 50 images per month
password = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';

var options = {
  'method' : 'post',
  'payload' : formData
};
// Replace username with your User ID and password with your API Key
options.headers = {"Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)};
if (sent != "sent") {
var res = UrlFetchApp.fetch("https://hcti.io/v1/image", options);
//Logger.log(res.getContentText());
var dataAll = JSON.parse(res.getContentText());
Logger.log(dataAll['url']) 
}

//////////////////////////////Sent Image from API To line notify Section////////////////////////////////////////

 
  var Token1 = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';   Insert your line token

  var message = "ข้อมูล ประจำเดือน "+getcurrentmonth()+" ::";
  
  if(message!=""){
    if (sent != "sent") {
   sendMS(Token1,message); 
    }
  } 

function sendMS(Token,message)
{
var imgThumbnail = dataAll['url']; // Maximum size of 240×240px JPEG
var imgFullsize = dataAll['url']; //Maximum size of 1024×1024px JPEG

var formData = {
'message' : message,
'imageThumbnail': imgThumbnail,
'imageFullsize' : imgFullsize,
}

 
  var options = {
    'method' : 'post',
    'headers' : {'Authorization': "Bearer "+Token},
    'contentType': 'application/x-www-form-urlencoded',
    'payload' : formData
  };
   
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', options);
  
}


//Not use now
function getcurrentmonth() {
    var date = new Date();
    var mt = date.getMonth();
    var yt = date.getFullYear()+543;
    var months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤจิกายน", "ธันวาคม"];

    var currentD = months[mt] + " " + yt; 

    return currentD; 
}
}

//create menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send to Line')
    .addItem('Send', 'myFunction')
    .addToUi();
}

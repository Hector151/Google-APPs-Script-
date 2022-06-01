# Google-APPs-Script-
---
function myFunction() 
{
  ouputExcel_ket = '1WWlSaTgRRCyaOTj9LLQ_s2CimwLMgzUSwYSK_hu6S5w'
  inputFile_key = '1YXtXzGfuRfcD0MLpvYjZxBA2CZ820Z_T'
  var app = SpreadsheetApp.openById(ouputExcel_ket);// 想輸出的試算表金鑰
  var sheet = app.getSheets()[0];//EXCEL第一表格
  var range = sheet.getDataRange();//取得表格內容
  var values = range.getValues();//取得內容中的資料
  // var sheet2 = app.getSheets()[1];//EXCEL第二表格
  // var range2 = sheet2.getDataRange();//取得表格內容
  // var values2 = range2.getValues();//取得內容中的資料
  var videoFolder = DriveApp.getFolderById(inputFile_key);//放影片資料夾的資料夾金鑰
  var foldersInVideoFolder = videoFolder.getFolders();//獲取目錄中所資料夾的集合
  var folder;
  var folderID =[];//影片資料夾金鑰

  var ImageFormat=["jpg","jpeg","png","tif","tiff","bmp"]//圖片格式
  var DataFormat=["pptx", "pdf"] //資料格式
  var Videoformat=["wmv","mp4","rmvb"]

  for (var i = 0; foldersInVideoFolder.hasNext(); i++)
  {
      folder = foldersInVideoFolder.next();//next()獲取文件或文件夾集合中的下一項。
      folderID[i] = folder.getId();//在 folderID[]中存影片資料夾金鑰
      //var data =[folder.getName(),folder.getId()]
      // sheet.appendRow(data);
  }
  for (var i = 0; i < folderID.length; i++)
  {
    var folder = DriveApp.getFolderById(folderID[i]);//folder讀取folderID[]中存影片資料夾金鑰
    var files = folder.getFiles();   //getFiles()獲取目錄中所有文件的集合
    var file; 
    var DataQuantity=0,run=0;
    for(var j=0;files.hasNext();j++)
    {
      file=files.next();
      DataQuantity++;
      dotSplit = file.getName().split("."); //dotSplit取得文件名稱，並以"."來分割
      var count=0; count2=0, count3=0//count圖片格式，count2影片格式,count3是資料格式
      while((count3<DataFormat.length))
      {
        if(dotSplit[dotSplit.length-1]==DataFormat[count3])
        {//判斷副檔名是否為影片格式
          var data =[folder.getName(),folder.getId(),file.getName(),file.getUrl()]
          sheet.appendRow(data);
          count3=DataFormat.length;//成功輸出過就暫停
        }
        else
        {
          count3++,run++
          while(count2<Videoformat.length) 
          {
            if(dotSplit[dotSplit.length-1]==Videoformat[count2])  
            {//判斷是否為影片
              var data =[folder.getName(),file.getId(),file.getName(),"https://drive.google.com/file/d/"+file.getId()+"/preview"]
              // sheet2.appendRow(data);
              sheet.appendRow(data)
              count2=Videoformat.length;
            }
            else  {count2++;}
          }
        }   //繼續判斷副檔名
      }
    }
    // if(DataQuantity==run/DataFormat.length){//如果資料夾中沒有圖片，輸出資料夾名稱跟金鑰
    //     var data =[folder.getName(),folder.getId()]
    //     sheet.appendRow(data);
    //     }
  }
  sheet.sort(1) //有片排序
}

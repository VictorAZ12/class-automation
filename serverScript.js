function doGet(request) {
    // initiate html service
    return HtmlService.createTemplateFromFile('Index')
        .evaluate();
  }
  function include(filename) {
    // include css and js into html template (printing scriptlets)
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
  }
  
  function retrieveFolders(){
    // retrieve all folders and display names
    let folders = DriveApp.getFolders();
    while (folders.hasNext()) {
      let folder = folders.next();
      Logger.log("Folder: %s, id %s",folder.getName(), folder.getId());
    }
  }
  
  function getFoldersInFolder(folderId) {
    // retrieve the given folder, otherwise root folder is used
    // return a list of folder in JSON: {[{id, name}, ...]}
  
    // get folder
    let folder;
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    } else {
      folder = DriveApp.getRootFolder();
    }
    
    // get subfolders
    let folders = folder.getFolders();
    let folderList = [];
    while (folders.hasNext()) {
      let subFolder = folders.next();
      folderList.push({
        id: subFolder.getId(),
        name: subFolder.getName()
      });
    }
    return JSON.stringify(folderList);
  }
  
  function getFolderPath(folderId) {
    // get a folder's path structure using its id
    // return a list folder path in JSON {[{id, name}, ...]}
    // the first element should be the root folder "Drive App", the last element should be the 
    // given folder if it's not root
  
    // get folder
    let path = [];
    let folder;
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    }
    else {
      folder = DriveApp.getRootFolder();
    }
    
    // get all parent folders
    while (folder) {
      path.unshift({
        id: folder.getId(),
        name: folder.getName()
      });
  
      let parents = folder.getParents();
      if (parents.hasNext()) {
        folder = parents.next();
      }
      else {
        folder = null;
      }
    }
    return JSON.stringify(path);
    
  }
  
  function getFolderPathString(folderId){
    // construct a path string "Drive App/.../.../folder"
    return JSON.parse(getFolderPath(folderId)).map(folder => folder.name).join("/");
  }
  function extractSheetId(url) {
  var match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}
    function processSheet(sheetUrl) {
    // sheetUrl = 'https://docs.google.com/spreadsheets/d/1gGT-JXSgSW29eIbVNaTj_QuQEh-_tVa8Z-fVD3qjY_k';
    try {
    // 提取 Google Sheet 文件 ID
    var sheetId = extractSheetId(sheetUrl);
    if (!sheetId) throw new Error('Invalid Google Sheet URL.');

    // 打开 Google Sheet
    var spreadsheet = SpreadsheetApp.openById(sheetId);

    // 获取 "Course Data" 工作表
    var courseSheet = spreadsheet.getSheetByName('Course Data');
    if (!courseSheet) throw new Error('Sheet "Course Data" not found.');

    // 获取 "Teacher and Room data" 工作表
    var teacherSheet = spreadsheet.getSheetByName('Teacher and Room data');
    if (!teacherSheet) throw new Error('Sheet "Teacher and Room data" not found.');

    // 读取 "Teacher and Room data" 的 Teacher 和 Email 列
    var teacherData = teacherSheet.getRange('A2:B' + teacherSheet.getLastRow()).getValues();
    var teacherEmailMap = {};
    teacherData.forEach(row => {
      if (row[0] && row[1]) {
        teacherEmailMap[row[0].toString().trim()] = row[1].toString().trim();
      }
    });

    // 读取 "Course Data" 的所有数据
    var courseData = courseSheet.getDataRange().getValues();
    if (courseData.length <= 1) throw new Error('No data found in "Course Data" sheet.');

    // 提取 Weekday 不为空的数据行（除了第一行）
    var headers = courseData[0]; // 第一行是表头
    var weekdayIndex = headers.indexOf('Weekday');
    if (weekdayIndex === -1) throw new Error('Column "Weekday" not found.');

    var filteredData = [];
    for (var i = 1; i < courseData.length; i++) {
      if (courseData[i][weekdayIndex] && courseData[i][weekdayIndex].toString().trim() !== '') {
        filteredData.push(courseData[i]);
      }
    }
    if (filteredData.length === 0) throw new Error('No rows with non-empty Weekday found.');

    // 处理老师课表
    var teacherSchedules = {};
    var teacherIndex = headers.indexOf('Teacher');
    if (teacherIndex === -1) throw new Error('Column "Teacher" not found.');

    filteredData.forEach(row => {
      // 分割 Teacher 列（可能包含多个老师，用逗号分隔）
      var teachers = row[teacherIndex].toString().split(',').map(t => t.trim());
      teachers.forEach(teacher => {
        if (teacher && teacherEmailMap[teacher]) {
          if (!teacherSchedules[teacher]) {
            teacherSchedules[teacher] = {
              email: teacherEmailMap[teacher],
              classes: []
            };
          }
          // 将课程信息添加到老师的课表
          var classInfo = {};
          headers.forEach((header, index) => {
            classInfo[header] = row[index];
          });
          teacherSchedules[teacher].classes.push(classInfo);
        }
      });
    });

    // 转换为要求的格式：数组形式
    var result = Object.keys(teacherSchedules).map(teacher => ({
      [teacher]: teacherSchedules[teacher]
    }));
    Logger.log(JSON.stringify(result))
    // 返回结果
    return JSON.stringify({
      status: 'success',
      data: result
    });
  } catch (e) {
    throw new Error('Failed to process sheet: ' + e.message);
  }
}
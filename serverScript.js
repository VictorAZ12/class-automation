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

// 辅助函数：将 Date 对象格式化为 hh:mm 字符串
function formatTime(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  var hours = date.getHours().toString().padStart(2, '0');
  var minutes = date.getMinutes().toString().padStart(2, '0');
  return `${hours}:${minutes}`;
}

function processSheetData(folderId, sheetUrl) {
  folderId = '1xg9SjmOe_uIDKr8BzDwWpyPMhzyNWgXX';
  sheetUrl = 'https://docs.google.com/spreadsheets/d/1gGT-JXSgSW29eIbVNaTj_QuQEh-_tVa8Z-fVD3qjY_k'
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
    var startTimeIndex = headers.indexOf('Start Time (hh:mm)');
    var endTimeIndex = headers.indexOf('End Time (hh:mm)');
    if (startTimeIndex === -1 || endTimeIndex === -1) throw new Error('Time columns not found.');
    for (var i = 1; i < courseData.length; i++) {
      if (courseData[i][weekdayIndex] && courseData[i][weekdayIndex].toString().trim() !== '') {
        var row = courseData[i].slice(); // 复制行数据
        // 格式化时间字段为 hh:mm
        if (row[startTimeIndex] instanceof Date) {
          row[startTimeIndex] = formatTime(row[startTimeIndex]);
        }
        if (row[endTimeIndex] instanceof Date) {
          row[endTimeIndex] = formatTime(row[endTimeIndex]);
        }
        filteredData.push(row);
      }
    }
    if (filteredData.length === 0) throw new Error('No rows with non-empty Weekday found.');

    // 处理老师课表
    var teacherSchedules = {};
    var teacherIndex = headers.indexOf('Teacher');
    if (teacherIndex === -1) throw new Error('Column "Teacher" not found.');

    filteredData.forEach(row => {
      var teachers = row[teacherIndex].toString().split(',').map(t => t.trim());
      teachers.forEach(teacher => {
        if (teacher && teacherEmailMap[teacher]) {
          if (!teacherSchedules[teacher]) {
            teacherSchedules[teacher] = {
              email: teacherEmailMap[teacher],
              classes: []
            };
          }
          var classInfo = {};
          headers.forEach((header, index) => {
            classInfo[header] = row[index];
          });
          teacherSchedules[teacher].classes.push(classInfo);
        }
      });
    });
    Logger.log(JSON.stringify(teacherSchedules));
    // 获取输出文件夹
    var folder = DriveApp.getFolderById(folderId);

    // 为每位老师生成 Google Sheet 课表
    Object.keys(teacherSchedules).forEach(teacher => {
      var { email, classes } = teacherSchedules[teacher];
      if (classes.length === 0) return; // 跳过没有课程的老师

      // 确定时间范围
      var times = [];
      classes.forEach(cls => {
        var start = new Date(cls['Start Time (hh:mm)']);
        var end = new Date(cls['End Time (hh:mm)']);
        times.push(start, end);
      });

      var minTime = new Date(Math.min(...times));
      var maxTime = new Date(Math.max(...times));
      // 调整 minTime 到最近的 15 分钟间隔
      minTime.setSeconds(0, 0);
      minTime.setMinutes(Math.floor(minTime.getMinutes() / 15) * 15);
      // 调整 maxTime 到下一个 15 分钟间隔
      maxTime.setSeconds(0, 0);
      maxTime.setMinutes(Math.ceil(maxTime.getMinutes() / 15) * 15);

      // 生成时间槽
      var timeSlots = [];
      var current = new Date(minTime);
      while (current <= maxTime) {
        timeSlots.push(new Date(current));
        current.setMinutes(current.getMinutes() + 15);
      }

      // 初始化课表数据（行：时间槽，列：星期一到星期日）
      var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
      var scheduleData = timeSlots.map(() => days.map(() => ''));

      // 填充课表
      classes.forEach(cls => {
        var start = new Date(cls['Start Time (hh:mm)']);
        var end = new Date(cls['End Time (hh:mm)']);
        var day = cls['Weekday'];
        var dayIndex = days.indexOf(day);
        if (dayIndex === -1) return; // 跳过无效星期

        // 找到时间范围内的行
        var startRow = timeSlots.findIndex(t => t.getTime() >= start.getTime());
        var endRow = timeSlots.findIndex(t => t.getTime() >= end.getTime());
        if (startRow === -1 || endRow === -1) return;

        // 填充内容（按优先级）
        var contents = [
          `${formatTime(start)} - ${formatTime(end)}`,
          cls['Course Name'] || '',
          cls['Room'] || '',
          cls['Course Type'] || '',
          cls['Notes'] || ''
        ];
        for (var i = startRow; i <= endRow && i < timeSlots.length; i++) {
          if (i - startRow < contents.length) {
            scheduleData[i][dayIndex] = contents[i - startRow];
          }
        }
      });

      // 创建新的 Google Sheet
      var newSpreadsheet = SpreadsheetApp.create(`${teacher}'s Weekly Schedule`);
      var sheet = newSpreadsheet.getSheets()[0];

      // 设置表头（第一列为空，第二列到第八列为星期）
      var headerRow = ['Time', ...days];
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

      // 设置时间列和课表数据
      var timeLabels = timeSlots.map(t => formatTime(t));
      var dataRange = sheet.getRange(2, 1, timeSlots.length, headerRow.length);
      var dataValues = timeSlots.map((_, i) => [timeLabels[i], ...scheduleData[i]]);
      dataRange.setValues(dataValues);

      // 格式化表格
      sheet.getRange(1, 1, 1, headerRow.length).setFontWeight('bold');
      sheet.getRange(2, 1, timeSlots.length, 1).setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);

      // 移动文件到指定文件夹
      var file = DriveApp.getFileById(newSpreadsheet.getId());
      file.moveTo(folder);

      // 分享文件给老师
      try {
        file.addEditor(email);
      } catch (e) {
        Logger.log(`Failed to share file with ${email}: ${e.message}`);
      }
    });

    return JSON.stringify({
      status: 'success',
      message: `Created ${Object.keys(teacherSchedules).length} teacher schedules in the specified folder.`
    });
  } catch (e) {
    throw new Error('Failed to process sheet: ' + e.message);
  }
}


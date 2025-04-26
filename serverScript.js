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

// 辅助函数：提取 Google Sheet ID
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

// 辅助函数：解析 hh:mm 字符串为当天的 Date 对象
function parseTime(timeStr) {
if (!timeStr || typeof timeStr !== 'string') return null;
var [hours, minutes] = timeStr.split(':').map(Number);
if (isNaN(hours) || isNaN(minutes)) return null;
var date = new Date();
date.setHours(hours, minutes, 0, 0);
return date;
}

/**
* 处理多个 Google Sheets 的数据，生成教师课表
* @param {string} folderId 输出文件夹的 Google Drive ID
* @param {Object[]} sheetData 包含 [{ sheetLink, campusName }, ...] 的数组
* @returns {string} JSON 格式的结果，包含处理状态和消息
*/
function processSheetData(folderId, sheetData) {
try {
  // 验证输入
  if (!folderId) throw new Error('Output folder ID is missing.');
  if (!sheetData || !Array.isArray(sheetData) || sheetData.length === 0) {
    throw new Error('Sheet data is empty or invalid.');
  }

  // 获取输出文件夹
  var folder = DriveApp.getFolderById(folderId);

  // 全局教师映射：Email 到首次出现的教师名字
  var emailToTeacherMap = {};
  // 校区特定的教师到 Email 映射：{ campusName: { teacherName: email } }
  var teacherToEmailMap = {};
  // 教师课表数据：{ email: { classes: [], campusNames: Set } }
  var teacherSchedules = {};

  // 处理结果
  var results = [];

  // 遍历每个 Sheet
  sheetData.forEach(({ sheetLink, campusName }, index) => {
    try {
      Logger.log(`Processing sheet ${index + 1}: ${sheetLink} (Campus: ${campusName})`);

      // 提取 Sheet ID
      var sheetId = extractSheetId(sheetLink);
      if (!sheetId) throw new Error('Invalid Google Sheet URL.');

      // 打开 Google Sheet
      var spreadsheet = SpreadsheetApp.openById(sheetId);

      // 检查 Control 表
      var controlSheet = spreadsheet.getSheetByName('Control');
      if (!controlSheet) throw new Error('Sheet "Control" not found.');
      var controlData = controlSheet.getRange('A1:B2').getValues();
      if (controlData[0][0] !== 'Flag' || controlData[0][1] !== 'Value') {
        throw new Error('Invalid headers in "Control" sheet.');
      }
      if (controlData[1][0] !== 'generateTeacherSchedule?' || controlData[1][1].toString().toLowerCase() !== 'yes') {
        Logger.log(`Skipping ${campusName}: generateTeacherSchedule? is not "Yes".`);
        results.push({
          campusName,
          status: 'skipped',
          message: 'generateTeacherSchedule? is not set to Yes.'
        });
        return;
      }

      // 获取 Teacher Data 表
      var teacherSheet = spreadsheet.getSheetByName('Teacher Data');
      if (!teacherSheet) throw new Error('Sheet "Teacher Data" not found.');
      var teacherData = teacherSheet.getRange('A1:C' + teacherSheet.getLastRow()).getValues();
      if (teacherData[0][0] !== 'Teacher' || teacherData[0][1] !== 'Email' || teacherData[0][2] !== 'NeedUpdate?') {
        throw new Error('Invalid headers in "Teacher Data" sheet.');
      }

      // 构建校区教师到 Email 的映射
      teacherToEmailMap[campusName] = {};
      for (var i = 1; i < teacherData.length; i++) {
        var teacherName = teacherData[i][0] ? teacherData[i][0].toString().trim() : '';
        var email = teacherData[i][1] ? teacherData[i][1].toString().trim() : '';
        if (teacherName && email) {
          teacherToEmailMap[campusName][teacherName] = email;
          if (!emailToTeacherMap[email]) {
            emailToTeacherMap[email] = teacherName; // 记录首次出现的教师名字
          }
        }
      }

      // 获取 Course Data 表
      var courseSheet = spreadsheet.getSheetByName('Course Data');
      if (!courseSheet) throw new Error('Sheet "Course Data" not found.');
      var courseData = courseSheet.getDataRange().getValues();
      if (courseData.length <= 1) throw new Error('No data found in "Course Data" sheet.');

      // 验证表头
      var headers = courseData[0];
      var requiredHeaders = [
        'Validation', 'Weekday', 'Start Time', 'End Time', 'Room',
        'Teacher', 'Course Name', 'Course Type', 'Student Count', 'Students', 'Notes'
      ];
      var missingHeaders = requiredHeaders.filter(h => headers.indexOf(h) === -1);
      if (missingHeaders.length > 0) {
        throw new Error(`Missing headers in "Course Data": ${missingHeaders.join(', ')}`);
      }

      // 提取有效课程数据
      var weekdayIndex = headers.indexOf('Weekday');
      var startTimeIndex = headers.indexOf('Start Time');
      var endTimeIndex = headers.indexOf('End Time');
      var teacherIndex = headers.indexOf('Teacher');

      var filteredData = [];
      for (var i = 1; i < courseData.length; i++) {
        if (courseData[i][weekdayIndex] && courseData[i][weekdayIndex].toString().trim() !== '') {
          var row = courseData[i].slice();
          // 格式化时间字段为 hh:mm
          if (row[startTimeIndex] instanceof Date) {
            row[startTimeIndex] = formatTime(row[startTimeIndex]);
          }
          if (row[endTimeIndex] instanceof Date) {
            row[endTimeIndex] = formatTime(row[endTimeIndex]);
          }
          // 验证时间
          var start = parseTime(row[startTimeIndex]);
          var end = parseTime(row[endTimeIndex]);
          if (!start || !end) {
            Logger.log(`Warning: Invalid time format in ${campusName}, row ${i + 1}: Start=${row[startTimeIndex]}, End=${row[endTimeIndex]}`);
            continue;
          }
          if (end.getTime() <= start.getTime()) {
            Logger.log(`Warning: End time before start time in ${campusName}, row ${i + 1}: Start=${row[startTimeIndex]}, End=${row[endTimeIndex]}`);
            continue;
          }
          // 检查异常长的课程（超过 3 小时）
          var durationHours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
          if (durationHours > 3) {
            Logger.log(`Warning: Unusually long course in ${campusName}, row ${i + 1}: ${durationHours.toFixed(2)} hours`);
          }
          filteredData.push(row);
        }
      }

      // 处理课程数据
      filteredData.forEach(row => {
        var teachers = row[teacherIndex].toString().split(',').map(t => t.trim());
        teachers.forEach(teacher => {
          if (teacher && teacherToEmailMap[campusName][teacher]) {
            var email = teacherToEmailMap[campusName][teacher];
            if (!teacherSchedules[email]) {
              teacherSchedules[email] = {
                classes: [],
                campusNames: new Set()
              };
            }
            var classInfo = {};
            headers.forEach((header, index) => {
              classInfo[header] = row[index];
            });
            classInfo['Campus Name'] = campusName; // 添加校区名称
            teacherSchedules[email].classes.push(classInfo);
            teacherSchedules[email].campusNames.add(campusName);
          }
        });
      });

      results.push({
        campusName,
        status: 'success',
        message: `Processed ${filteredData.length} valid courses for ${campusName}.`
      });
    } catch (e) {
      Logger.log(`Error processing ${campusName}: ${e.message}`);
      results.push({
        campusName,
        status: 'error',
        message: `Failed to process ${campusName}: ${e.message}`
      });
    }
  });

  // 生成教师课表
  var colors = ['#f6d7b0', '#b7e1cd', '#b3cde3', '#f4c7c3']; // light orange, green, blue, red
  var colorIndex = 0;

  Object.keys(teacherSchedules).forEach(email => {
    var { classes, campusNames } = teacherSchedules[email];
    if (classes.length === 0) return;

    // 使用首次出现的教师名字（仅用于日志，文件名不再依赖 teacherName）
    var teacherName = emailToTeacherMap[email] || 'Unknown';
    Logger.log(`Generating schedule for ${teacherName} (${email})`);

    // 确定时间范围
    var times = [];
    classes.forEach(cls => {
      var start = parseTime(cls['Start Time']);
      var end = parseTime(cls['End Time']);
      if (start && end) {
        times.push(start, end);
      }
    });

    if (times.length === 0) {
      Logger.log(`No valid times for ${teacherName} (${email}). Skipping.`);
      return;
    }

    var minTime = new Date(Math.min(...times));
    var maxTime = new Date(Math.max(...times));
    minTime.setSeconds(0, 0);
    minTime.setMinutes(Math.floor(minTime.getMinutes() / 15) * 15);
    maxTime.setSeconds(0, 0);
    maxTime.setMinutes(Math.ceil(maxTime.getMinutes() / 15) * 15);

    // 生成时间槽
    var timeSlots = [];
    var current = new Date(minTime);
    while (current <= maxTime) {
      timeSlots.push(new Date(current));
      current.setMinutes(current.getMinutes() + 15);
    }

    // 初始化课表数据
    var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    var scheduleData = timeSlots.map(() => days.map(() => ''));
    var backgroundColors = timeSlots.map(() => days.map(() => null));

    // 填充课表
    classes.forEach(cls => {
      var start = parseTime(cls['Start Time']);
      var end = parseTime(cls['End Time']);
      if (!start || !end) return;
      var day = cls['Weekday'];
      var dayIndex = days.indexOf(day);
      if (dayIndex === -1) return;

      var startRow = timeSlots.findIndex(t => t.getTime() >= start.getTime());
      var endRow = timeSlots.findIndex(t => t.getTime() >= end.getTime());
      if (startRow === -1 || endRow === -1) return;

      var currentColor = colors[colorIndex % colors.length];
      colorIndex++;

      // 填充内容，首行包含校区名称
      var contents = [
        `${cls['Campus Name']}:[${formatTime(start)} - ${formatTime(end)}]`,
        cls['Course Name'] || '',
        cls['Room'] || '',
        cls['Course Type'] || '',
        cls['Notes'] || ''
      ];
      for (var i = startRow; i <= endRow && i < timeSlots.length; i++) {
        if (i - startRow < contents.length) {
          scheduleData[i][dayIndex] = contents[i - startRow];
        }
        backgroundColors[i][dayIndex] = currentColor;
      }
    });

    // 检查文件夹中是否已存在包含该 email 的课表文件
    var fileNamePrefix = `${email}'s Weekly Schedule`;
    var iterator = folder.getFiles();
    while (iterator.hasNext()) {
      var file = iterator.next();
      var fileName = file.getName();
      if (fileName.includes(fileNamePrefix)) {
        Logger.log(`Found existing file for ${email}: ${fileName}. Deleting.`);
        file.setTrashed(true); // 删除旧文件
      }
    }

    // 生成当前时间戳，格式为 YYYY-MM-DD HH:MM
    var now = new Date();
    var timeStamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var fileName = `${email}'s Weekly Schedule (${timeStamp})`;

    // 创建新的 Google Sheet
    var newSpreadsheet = SpreadsheetApp.create(fileName);
    var sheet = newSpreadsheet.getSheets()[0];

    // 设置表头
    var headerRow = ['Time', ...days];
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

    // 设置时间列和课表数据
    var timeLabels = timeSlots.map(t => formatTime(t));
    var dataRange = sheet.getRange(2, 1, timeSlots.length, headerRow.length);
    var dataValues = timeSlots.map((_, i) => [timeLabels[i], ...scheduleData[i]]);
    dataRange.setValues(dataValues);

    // 设置背景颜色
    for (var i = 0; i < timeSlots.length; i++) {
      for (var j = 0; j < days.length; j++) {
        if (backgroundColors[i][j]) {
          sheet.getRange(i + 2, j + 2).setBackground(backgroundColors[i][j]);
        }
      }
    }

    // 格式化表格
    sheet.getRange(1, 1, 1, headerRow.length).setFontWeight('bold');
    sheet.getRange(2, 1, timeSlots.length, 1).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
    sheet.getRange(1, 1, timeSlots.length + 1, headerRow.length).setHorizontalAlignment('center');
    for (var col = 1; col <= headerRow.length; col++) {
      sheet.setColumnWidth(col, 240);
    }

    // 移动文件到指定文件夹
    var file = DriveApp.getFileById(newSpreadsheet.getId());
    file.moveTo(folder);

    // 分享文件给老师（仅查看权限）
    try {
      file.addViewer(email);
      Logger.log(`Shared file with ${email} as viewer.`);
    } catch (e) {
      Logger.log(`Failed to share file with ${email}: ${e.message}`);
    }
  });

  // 汇总结果
  var successCount = results.filter(r => r.status === 'success').length;
  var errorCount = results.filter(r => r.status === 'error').length;
  var skippedCount = results.filter(r => r.status === 'skipped').length;
  var message = `Processed ${sheetData.length} sheets: ${successCount} succeeded, ${errorCount} failed, ${skippedCount} skipped.`;
  Logger.log(message);

  return JSON.stringify({
    status: 'success',
    message,
    details: results
  });
} catch (e) {
  Logger.log(`Fatal error: ${e.message}`);
  return JSON.stringify({
    status: 'error',
    message: `Failed to process sheets: ${e.message}`,
    details: []
  });
}
}
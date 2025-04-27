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
  if (!folderId) throw new Error('输出文件夹 ID 缺失。');
  if (!sheetData || !Array.isArray(sheetData) || sheetData.length === 0) {
    throw new Error('表格数据为空或格式无效。');
  }

  // 获取输出文件夹
  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`无法访问文件夹 ID ${folderId}：${e.message}`);
  }

  // 全局教师映射：Email 到首次出现的教师名字
  var emailToTeacherMap = {};
  // 校区特定的教师到 Email 映射：{ campusName: { teacherName: email } }
  var teacherToEmailMap = {};
  // 教师课表数据：{ email: { classes: [], campusNames: Set } }
  var teacherSchedules = {};
  // 教师 NeedUpdate? 状态：{ email: boolean }
  var teacherNeedUpdate = {};

  // 处理结果
  var results = [];

  // 遍历每个 Sheet
  sheetData.forEach(({ sheetLink, campusName }, index) => {
    try {
      Logger.log(`处理表格 ${index + 1}：${sheetLink}（校区：${campusName}）`);

      // 提取 Sheet ID
      var sheetId = extractSheetId(sheetLink);
      if (!sheetId) {
        results.push({
          campusName,
          status: 'error',
          message: `校区 ${campusName} 的 Google Sheet 链接无效。`
        });
        return;
      }

      // 打开 Google Sheet
      var spreadsheet = SpreadsheetApp.openById(sheetId);

      // 检查 Control 表
      var controlSheet = spreadsheet.getSheetByName('Control');
      if (!controlSheet) {
        results.push({
          campusName,
          status: 'error',
          message: `校区 ${campusName} 缺少 "Control" 表。`
        });
        return;
      }
      var controlData = controlSheet.getRange('A1:B2').getValues();
      if (controlData[0][0] !== 'Flag' || controlData[0][1] !== 'Value') {
        results.push({
          campusName,
          status: 'error',
          message: `校区 ${campusName} 的 "Control" 表表头无效。`
        });
        return;
      }
      if (controlData[1][0] !== 'generateTeacherSchedule?' || controlData[1][1].toString().toLowerCase() !== 'yes') {
        Logger.log(`跳过 ${campusName}：generateTeacherSchedule? 不是 "Yes"。`);
        results.push({
          campusName,
          status: 'skipped',
          message: `因为 Control flag 不是 Yes，所以跳过校区 ${campusName} 的处理。`
        });
        return;
      }

      // 获取 Teacher Data 表
      var teacherSheet = spreadsheet.getSheetByName('Teacher Data');
      if (!teacherSheet) {
        results.push({
          campusName,
          status: 'error',
          message: `校区 ${campusName} 缺少 "Teacher Data" 表。`
        });
        return;
      }

      // 获取表头行（第一行）
      var lastColumn = teacherSheet.getLastColumn();
      var headers = teacherSheet.getRange(1, 1, 1, lastColumn).getValues()[0];

      // 查找所需表头的列索引
      var teacherIndex = headers.indexOf('Teacher');
      var emailIndex = headers.indexOf('Email');
      var needUpdateIndex = headers.indexOf('NeedUpdate?');
      if (teacherIndex === -1 || emailIndex === -1 || needUpdateIndex === -1) {
        results.push({
          campusName,
          status: 'error',
          message: `校区 ${campusName} 的 "Teacher Data" 表缺少必要表头：Teacher, Email, NeedUpdate?`
        });
        return;
      }

      // 获取数据（从第 2 行到最后一行，仅读取需要的列）
      var lastRow = teacherSheet.getLastRow();
      var teacherData = [];
      if (lastRow > 1) { // 确保有数据行
        var teacherCol = teacherSheet.getRange(2, teacherIndex + 1, lastRow - 1, 1).getValues().map(row => row[0]);
        var emailCol = teacherSheet.getRange(2, emailIndex + 1, lastRow - 1, 1).getValues().map(row => row[0]);
        var needUpdateCol = teacherSheet.getRange(2, needUpdateIndex + 1, lastRow - 1, 1).getValues().map(row => row[0]);
        for (var i = 0; i < teacherCol.length; i++) {
          teacherData.push([teacherCol[i], emailCol[i], needUpdateCol[i]]);
        }
      }

      // 构建校区教师到 Email 的映射，并记录 NeedUpdate?
      teacherToEmailMap[campusName] = {};
      for (var i = 0; i < teacherData.length; i++) {
        var teacherName = teacherData[i][0] ? teacherData[i][0].toString().trim() : '';
        var email = teacherData[i][1] ? teacherData[i][1].toString().trim() : '';
        var needUpdate = teacherData[i][2] ? teacherData[i][2].toString().trim().toLowerCase() === 'yes' : false;
        if (teacherName && email) {
          teacherToEmailMap[campusName][teacherName] = email;
          if (emailToTeacherMap[email] && emailToTeacherMap[email] !== teacherName) {
            Logger.log(`警告：邮箱 ${email} 在校区 ${campusName} 对应多个教师名称：${emailToTeacherMap[email]} 和 ${teacherName}`);
            results.push({
              campusName,
              status: 'warning',
              message: `邮箱 ${email} 在校区 ${campusName} 对应多个教师名称：${emailToTeacherMap[email]} 和 ${teacherName}`
            });
          }
          if (!emailToTeacherMap[email]) {
            emailToTeacherMap[email] = teacherName; // 记录首次出现的教师名字
          }
          teacherNeedUpdate[email] = teacherNeedUpdate[email] || needUpdate; // 记录 NeedUpdate? 状态
        }
      }

      // 获取 Course Data 表
      var courseSheet = spreadsheet.getSheetByName('Course Data');
      if (!courseSheet) {
        results.push({
          campusName,
          status: 'error',
          message: `校区 ${campusName} 缺少 "Course Data" 表。`
        });
        return;
      }
      var courseData = courseSheet.getDataRange().getValues();
      if (courseData.length <= 1) {
        results.push({
          campusName,
          status: 'error',
          message: `校区 ${campusName} 的 "Course Data" 表没有数据。`
        });
        return;
      }

      // 验证表头
      var headers = courseData[0];
      var requiredHeaders = [
        'Validation', 'Weekday', 'Start Time', 'End Time', 'Room',
        'Teacher', 'Course Name', 'Course Type', 'Student Count', 'Students', 'Notes'
      ];
      var missingHeaders = requiredHeaders.filter(h => headers.indexOf(h) === -1);
      if (missingHeaders.length > 0) {
        results.push({
          campusName,
          status: 'error',
          message: `校区 ${campusName} 的 "Course Data" 表缺少必要表头：${missingHeaders.join(', ')}`
        });
        return;
      }

      // 提取有效课程数据
      var weekdayIndex = headers.indexOf('Weekday');
      var startTimeIndex = headers.indexOf('Start Time');
      var endTimeIndex = headers.indexOf('End Time');
      var teacherIndex = headers.indexOf('Teacher');

      var validDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
      var filteredData = [];
      for (var i = 1; i < courseData.length; i++) {
        var weekday = courseData[i][weekdayIndex] ? courseData[i][weekdayIndex].toString().trim() : '';
        // 验证 Weekday
        if (!weekday || !validDays.includes(weekday)) {
          results.push({
            campusName,
            status: 'warning',
            message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行 weekday 无效：${weekday}，该课程未加入课表。`
          });
          Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行 weekday 无效：${weekday}`);
          continue;
        }
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
          results.push({
            campusName,
            status: 'warning',
            message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行时间格式无效：开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}，该课程未加入课表。`
          });
          Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行时间格式无效：开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}`);
          continue;
        }
        if (end.getTime() <= start.getTime()) {
          results.push({
            campusName,
            status: 'warning',
            message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行结束时间早于或等于开始时间：开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}，该课程未加入课表。`
          });
          Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行结束时间早于或等于开始时间：开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}`);
          continue;
        }
        // 检查异常长的课程（超过 3 小时）
        var durationHours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
        if (durationHours > 3) {
          Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行课程时间过长：${durationHours.toFixed(2)} 小时`);
          results.push({
            campusName,
            status: 'warning',
            message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行课程时间过长：${durationHours.toFixed(2)} 小时`
          });
        }
        filteredData.push(row);
      }

      // 处理课程数据
      filteredData.forEach((row, rowIndex) => {
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
          } else if (teacher) {
            results.push({
              campusName,
              status: 'warning',
              message: `校区 ${campusName} 的 "Course Data" 表第 ${rowIndex + 2} 行教师 ${teacher} 未在 "Teacher Data" 表中找到，未加入课表。`
            });
            Logger.log(`警告：校区 ${campusName} 的第 ${rowIndex + 2} 行教师 ${teacher} 未在 "Teacher Data" 表中找到`);
          }
        });
      });

      results.push({
        campusName,
        status: 'success',
        message: `成功处理校区 ${campusName} 的 ${filteredData.length} 门有效课程。`
      });
    } catch (e) {
      Logger.log(`处理校区 ${campusName} 出错：${e.message}`);
      results.push({
        campusName,
        status: 'error',
        message: `处理校区 ${campusName}  Gabriella 失败：${e.message}`
      });
    }
  });

  // 检测时间冲突并标记有冲突的教师
  var teacherConflicts = {};
  Object.keys(teacherSchedules).forEach(email => {
    var { classes } = teacherSchedules[email];
    teacherConflicts[email] = { hasConflict: false, conflictDetails: [] };

    // 按 Weekday 分组课程
    var classesByDay = {};
    classes.forEach(cls => {
      var day = cls['Weekday'];
      if (!classesByDay[day]) classesByDay[day] = [];
      classesByDay[day].push(cls);
    });

    // 检查每个 Weekday 的冲突
    Object.keys(classesByDay).forEach(day => {
      var dayClasses = classesByDay[day];
      // 按开始时间排序，便于检查
      dayClasses.sort((a, b) => {
        var startA = parseTime(a['Start Time']);
        var startB = parseTime(b['Start Time']);
        return startA - startB;
      });

      // 两两比较课程
      for (var i = 0; i < dayClasses.length; i++) {
        for (var j = i + 1; j < dayClasses.length; j++) {
          var classA = dayClasses[i];
          var classB = dayClasses[j];
          var startA = parseTime(classA['Start Time']);
          var endA = parseTime(classA['End Time']);
          var startB = parseTime(classB['Start Time']);
          var endB = parseTime(classB['End Time']);

          if (!startA || !endA || !startB || !endB) continue;

          // 检查是否冲突（排除首尾相接）
          if (startA < endB && startB < endA && !(endA.getTime() === startB.getTime()) && !(endB.getTime() === startA.getTime())) {
            teacherConflicts[email].hasConflict = true;
            teacherConflicts[email].conflictDetails.push({
              day,
              classA: {
                campus: classA['Campus Name'],
                course: classA['Course Name'],
                time: `${classA['Start Time']}-${classA['End Time']}`
              },
              classB: {
                campus: classB['Campus Name'],
                course: classB['Course Name'],
                time: `${classB['Start Time']}-${classB['End Time']}`
              }
            });
          }
        }
      }
    });
  });

  // 生成教师课表
  var colors = ['#f6d7b0', '#b7e1cd', '#b3cde3', '#f4c7c3']; // light orange, green, blue, red
  var colorIndex = 0;

  Object.keys(teacherSchedules).forEach(email => {
    var teacherName = emailToTeacherMap[email] || '未知教师';
    // 检查时间冲突
    if (teacherConflicts[email].hasConflict) {
      var conflictMessages = teacherConflicts[email].conflictDetails.map(c =>
        `在 ${c.day}，${c.classA.campus} 的课程 ${c.classA.course}（${c.classA.time}）与 ${c.classB.campus} 的课程 ${c.classB.course}（${c.classB.time}）时间冲突`
      );
      Logger.log(`因时间冲突跳过为 ${email} 生成课表：${conflictMessages.join('；')}`);
      results.push({
        campusName: '教师课表',
        status: 'error',
        message: `无法为教师 ${email}（${teacherName}）生成课表：${conflictMessages.join('；')}`
      });
      return;
    }

    var { classes, campusNames } = teacherSchedules[email];
    // 检查是否需要更新
    if (!teacherNeedUpdate[email]) {
      Logger.log(`跳过为 ${email} 生成课表：NeedUpdate? 不是 "Yes"。`);
      results.push({
        campusName: '教师课表',
        status: 'skipped',
        message: `未为教师 ${email}（${teacherName}）生成课表：所有校区的 NeedUpdate? 均未设置为 Yes`
      });
      return;
    }
    // 检查空课程表
    if (classes.length === 0) {
      Logger.log(`为 ${email} 无有效课程，跳过生成课表。`);
      results.push({
        campusName: '教师课表',
        status: 'warning',
        message: `无法为教师 ${email}（${teacherName}）生成课表：未找到有效课程`
      });
      return;
    }

    Logger.log(`为 ${teacherName}（${email}）生成课表`);

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
      Logger.log(`为 ${teacherName}（${email}）无有效时间，跳过生成课表。`);
      results.push({
        campusName: '教师课表',
        status: 'warning',
        message: `无法为教师 ${email}（${teacherName}）生成课表：未找到有效课程时间`
      });
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
        Logger.log(`发现 ${email} 的现有课表文件：${fileName}，正在删除。`);
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
    var headerRow = ['时间', ...days.map(d => d === 'Monday' ? '星期一' : d === 'Tuesday' ? '星期二' : d === 'Wednesday' ? '星期三' : d === 'Thursday' ? '星期四' : d === 'Friday' ? '星期五' : d === 'Saturday' ? '星期六' : '星期日')];
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
      Logger.log(`已将课表分享给 ${email}（仅查看权限）。`);
      results.push({
        campusName: '教师课表',
        status: 'success',
        message: `为教师 ${email}（${teacherName}）成功生成课表`
      });
    } catch (e) {
      results.push({
        campusName: '教师课表',
        status: 'warning',
        message: `为教师 ${email}（${teacherName}）生成课表成功，但分享失败：${e.message}`
      });
      Logger.log(`无法将课表分享给 ${email}：${e.message}`);
    }
  });

  // 汇总结果
  var successCount = results.filter(r => r.status === 'success').length;
  var errorCount = results.filter(r => r.status === 'error').length;
  var skippedCount = results.filter(r => r.status === 'skipped').length;
  var warningCount = results.filter(r => r.status === 'warning').length;
  var message = `处理了 ${sheetData.length} 个校区表格：${successCount} 个成功，${errorCount} 个失败，${skippedCount} 个跳过，${warningCount} 个警告。`;
  Logger.log(message);

  return JSON.stringify({
    status: 'success',
    message,
    details: results
  });
} catch (e) {
  Logger.log(`严重错误：${e.message}`);
  return JSON.stringify({
    status: 'error',
    message: `处理表格失败：${e.message}`,
    details: results
  });
}
}
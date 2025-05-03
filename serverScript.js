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
 * 生成校区课表
 * @param {Spreadsheet} spreadsheet 校区 Google Sheet 对象
 * @param {string} campusName 校区名称
 * @param {string} folderId 输出文件夹的 Google Drive ID
 * @param {Object[]} results 结果数组，用于记录日志
 * @returns {void}
 */
function processGenerateCampusSchedule(spreadsheet, campusName, folderId, results) {
  // spreadsheet = SpreadsheetApp.openById('1SoBVlLIIrOkcrqvrLgwUzm5IH9RenEp_67F9dAdMINs');
  // campusName = 'CampusA';
  // folderId = '1xg9SjmOe_uIDKr8BzDwWpyPMhzyNWgXX';
  // results = [];
  try {
    Logger.log(`开始为校区 ${campusName} 生成课表`);
    var folder = DriveApp.getFolderById(folderId); // 获取目标文件夹

    // 获取 Course Data 表
    var courseSheet = spreadsheet.getSheetByName('Course Data');
    if (!courseSheet) {
      results.push({
        campusName,
        status: 'error',
        message: `校区 ${campusName} 缺少 "Course Data" 表，无法生成校区课表。`
      });
      Logger.log(`错误：校区 ${campusName} 缺少 "Course Data" 表`);
      return;
    }
    var courseData = courseSheet.getDataRange().getValues();
    if (courseData.length <= 1) {
      results.push({
        campusName,
        status: 'error',
        message: `校区 ${campusName} 的 "Course Data" 表没有数据，无法生成校区课表。`
      });
      Logger.log(`错误：校区 ${campusName} 的 "Course Data" 表没有数据`);
      return;
    }
    Logger.log(`校区 ${campusName} 的 Course Data 表有 ${courseData.length - 1} 行数据`);
    Logger.log(`校区 ${campusName} 的 Course Data 原始数据：${JSON.stringify(courseData)}`);

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
        message: `校区 ${campusName} 的 "Course Data" 表缺少必要表头：${missingHeaders.join(', ')}，无法生成校区课表。`
      });
      Logger.log(`错误：校区 ${campusName} 的 Course Data 表缺少表头：${missingHeaders}`);
      return;
    }
    Logger.log(`校区 ${campusName} 的 Course Data 表头：${headers}`);

    // 查找列索引
    var weekdayIndex = headers.indexOf('Weekday');
    var startTimeIndex = headers.indexOf('Start Time');
    var endTimeIndex = headers.indexOf('End Time');
    var roomIndex = headers.indexOf('Room');
    var courseNameIndex = headers.indexOf('Course Name');
    Logger.log(`校区 ${campusName} 的列索引：Weekday=${weekdayIndex}, Start Time=${startTimeIndex}, End Time=${endTimeIndex}, Room=${roomIndex}, Course Name=${courseNameIndex}`);
    if (weekdayIndex === -1 || startTimeIndex === -1 || endTimeIndex === -1 || roomIndex === -1 || courseNameIndex === -1) {
      results.push({
        campusName,
        status: 'error',
        message: `校区 ${campusName} 的 "Course Data" 表缺少必要列：Weekday=${weekdayIndex}, Start Time=${startTimeIndex}, End Time=${endTimeIndex}, Room=${roomIndex}, Course Name=${courseNameIndex}`
      });
      Logger.log(`错误：校区 ${campusName} 的列索引无效`);
      return;
    }

    // 提取有效课程数据
    var validDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    var filteredData = [];
    for (var i = 1; i < courseData.length; i++) {
      var row = courseData[i].slice();
      Logger.log(`校区 ${campusName} 的第 ${i + 1} 行原始数据：${JSON.stringify(row)}`);
      var weekday = row[weekdayIndex] ? row[weekdayIndex].toString().trim() : '';
      Logger.log(`校区 ${campusName} 的第 ${i + 1} 行原始 Weekday：${row[weekdayIndex]}, 处理后：${weekday}`);
      if (!weekday || !validDays.includes(weekday)) {
        results.push({
          campusName,
          status: 'warning',
          message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行 weekday 无效：${weekday}，该课程未加入校区课表。`
        });
        Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行 weekday 无效：${weekday}`);
        continue;
      }
      if (row[startTimeIndex] instanceof Date) {
        row[startTimeIndex] = formatTime(row[startTimeIndex]);
      }
      if (row[endTimeIndex] instanceof Date) {
        row[endTimeIndex] = formatTime(row[endTimeIndex]);
      }
      var start = parseTime(row[startTimeIndex]);
      var end = parseTime(row[endTimeIndex]);
      Logger.log(`校区 ${campusName} 的第 ${i + 1} 行时间：原始 Start=${row[startTimeIndex]}, End=${row[endTimeIndex]}, 解析后 Start=${start ? formatTime(start) : 'null'}, End=${end ? formatTime(end) : 'null'}`);
      if (!start || !end) {
        results.push({
          campusName,
          status: 'warning',
          message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行时间格式无效：开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}，该课程未加入校区课表。`
        });
        Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行时间格式无效`);
        continue;
      }
      if (end.getTime() <= start.getTime()) {
        results.push({
          campusName,
          status: 'warning',
          message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行结束时间早于或等于开始时间：开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}，该课程未加入校区课表。`
        });
        Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行结束时间早于或等于开始时间`);
        continue;
      }
      var durationHours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
      if (durationHours > 3) {
        results.push({
          campusName,
          status: 'warning',
          message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行课程时间过长：${durationHours.toFixed(2)} 小时`
        });
        Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行课程时间过长：${durationHours.toFixed(2)} 小时`);
      }
      var room = row[roomIndex] ? row[roomIndex].toString().trim() : '';
      Logger.log(`校区 ${campusName} 的第 ${i + 1} 行原始 Room：${row[roomIndex]}, 处理后：${room}`);
      if (!room) {
        results.push({
          campusName,
          status: 'warning',
          message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行 Room 为空，该课程未加入校区课表。`
        });
        Logger.log(`警告：校区 ${campusName} 的第 ${i + 1} 行 Room 为空`);
        continue;
      }
      filteredData.push(row);
      Logger.log(`校区 ${campusName} 的第 ${i + 1} 行有效课程：${JSON.stringify({
        Weekday: row[weekdayIndex],
        StartTime: row[startTimeIndex],
        EndTime: row[endTimeIndex],
        Room: row[roomIndex],
        CourseName: row[courseNameIndex]
      })}`);
    }
    Logger.log(`校区 ${campusName} 的 filteredData 包含 ${filteredData.length} 门有效课程`);

    if (filteredData.length === 0) {
      results.push({
        campusName,
        status: 'warning',
        message: `校区 ${campusName} 没有有效课程数据，无法生成校区课表。`
      });
      Logger.log(`警告：校区 ${campusName} 没有有效课程数据`);
      return;
    }

    // 获取所有教室
    var rooms = [...new Set(filteredData.map(row => row[roomIndex] ? row[roomIndex].toString().trim() : ''))].filter(r => r);
    if (rooms.length === 0) {
      results.push({
        campusName,
        status: 'error',
        message: `校区 ${campusName} 的 "Course Data" 表没有有效教室信息，无法生成校区课表。`
      });
      Logger.log(`错误：校区 ${campusName} 没有有效教室信息`);
      return;
    }
    Logger.log(`校区 ${campusName} 的教室列表：${rooms}`);

    // 确定时间范围
    var times = [];
    filteredData.forEach(row => {
      var start = parseTime(row[startTimeIndex]);
      var end = parseTime(row[endTimeIndex]);
      if (start && end) {
        times.push(start, end);
      }
    });
    if (times.length === 0) {
      results.push({
        campusName,
        status: 'error',
        message: `校区 ${campusName} 没有有效课程时间，无法生成校区课表。`
      });
      Logger.log(`错误：校区 ${campusName} 没有有效课程时间`);
      return;
    }
    var minTime = new Date(Math.min(...times));
    var maxTime = new Date(Math.max(...times));
    minTime.setSeconds(0, 0);
    minTime.setMinutes(Math.floor(minTime.getMinutes() / 15) * 15);
    maxTime.setSeconds(0, 0);
    maxTime.setMinutes(Math.ceil(maxTime.getMinutes() / 15) * 15);
    Logger.log(`校区 ${campusName} 的时间范围：${formatTime(minTime)} 到 ${formatTime(maxTime)}`);

    // 生成时间槽
    var timeSlots = [];
    var current = new Date(minTime);
    while (current <= maxTime) {
      timeSlots.push(new Date(current));
      current.setMinutes(current.getMinutes() + 15);
    }
    Logger.log(`校区 ${campusName} 的时间槽数量：${timeSlots.length}，从 ${formatTime(timeSlots[0])} 到 ${formatTime(timeSlots[timeSlots.length - 1])}`);

    // 初始化颜色
    var colors = ['#f6d7b0', '#b7e1cd', '#b3cde3', '#f4c7c3'];
    var colorIndex = 0;

    // 查找或创建 Google Sheet
    var fileNamePrefix = `${campusName}_Timetable_`;
    var existingFile = null;
    var spreadsheetToUse = null;
    var files = folder.getFiles();

    while (files.hasNext()) {
      var file = files.next();
      if (file.getName().startsWith(fileNamePrefix) && !file.isTrashed()) {
        existingFile = file;
        spreadsheetToUse = SpreadsheetApp.openById(existingFile.getId());
        Logger.log(`找到已存在的校区课表文件：${existingFile.getName()} (ID: ${existingFile.getId()})`);
        break; // 找到第一个匹配的就使用
      }
    }

    var now = new Date();
    var timeStamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var newFileName = `${fileNamePrefix}${timeStamp}`;
    var isExistingFile = false; // 标记是否是更新现有文件

    if (spreadsheetToUse) {
      isExistingFile = true;
      Logger.log(`准备更新现有表格 ${existingFile.getName()} 的内容`);
      var sheets = spreadsheetToUse.getSheets();

      // 保留第一个 sheet，删除其他的
      for (var i = sheets.length - 1; i > 0; i--) { // 从后往前删，避免索引问题
        Logger.log(`删除旧 sheet: ${sheets[i].getName()}`);
        spreadsheetToUse.deleteSheet(sheets[i]);
      }
      // 第一个 sheet 稍后会被重命名和清空
      if (sheets.length > 0) {
         Logger.log(`保留第一个旧 sheet: ${sheets[0].getName()}，稍后将重命名并清空`);
      } else {
         // 如果意外地没有 sheet 了（理论上不会发生，除非文件损坏），则创建一个临时的
         spreadsheetToUse.insertSheet("temp_placeholder");
         Logger.log("原文件没有 sheet，已创建临时 sheet");
      }

    } else {
      // 创建新的 Google Sheet
      spreadsheetToUse = SpreadsheetApp.create(newFileName);
      Logger.log(`创建新 Google Sheet：${newFileName}`);
    }

    // 为每一天创建一个 sheet 并填充数据
    var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    days.forEach((day, dayIndex) => {
      var sheet;
      if (dayIndex === 0) {
        // 获取第一个 sheet (无论是保留的旧 sheet 还是新创建的默认 sheet)
        sheet = spreadsheetToUse.getSheets()[0];
        sheet.setName(day); // 重命名
        sheet.clearContents(); // 清空内容
        sheet.clearFormats(); // 清空格式 (可选，但推荐)
        Logger.log(`重用/创建并清空第一个 sheet，命名为：${day}`);
      } else {
        // 检查是否已存在同名 sheet (不太可能，因为前面删除了，但作为保险)
        var existingDaySheet = spreadsheetToUse.getSheetByName(day);
        if (existingDaySheet) {
           sheet = existingDaySheet;
           sheet.clearContents();
           sheet.clearFormats();
           Logger.log(`找到并清空已存在的 sheet：${day}`);
        } else {
           sheet = spreadsheetToUse.insertSheet(day);
           Logger.log(`创建新 sheet：${day}`);
        }
      }

      var dayData = filteredData.filter(row => {
        var weekday = row[weekdayIndex] ? row[weekdayIndex].toString().trim() : '';
        return weekday === day;
      });
      Logger.log(`校区 ${campusName} 的 ${day} sheet 有 ${dayData.length} 门课程：${JSON.stringify(dayData.map(row => ({
        StartTime: row[startTimeIndex],
        EndTime: row[endTimeIndex],
        Room: row[roomIndex],
        CourseName: row[courseNameIndex]
      })))}`);

      // 初始化课表数据
      var scheduleData = timeSlots.map(() => Array(rooms.length).fill(''));
      var backgroundColors = timeSlots.map(() => Array(rooms.length).fill(null));

            // 填充课表
            dayData.forEach((row, rowIndex) => {
              Logger.log(`校区 ${campusName} 的 ${day} 第 ${rowIndex + 1} 门课程 原始数据：${JSON.stringify(row)}`);
              var start = parseTime(row[startTimeIndex]);
              var end = parseTime(row[endTimeIndex]);
              if (!start || !end) {
                Logger.log(`跳过校区 ${campusName} 的 ${day} 第 ${rowIndex + 1} 门课程：时间解析失败，开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}`);
                return;
              }
              // 使用正确的 roomIndex (列索引) 获取 room 值
              var room = row[roomIndex] ? row[roomIndex].toString().trim() : '';
              Logger.log(`校区 ${campusName} 的 ${day} 第 ${rowIndex + 1} 门课程：原始 Room 值=${row[roomIndex]}, 处理后=${room}`);
              if (!room) {
                Logger.log(`跳过校区 ${campusName} 的 ${day} 第 ${rowIndex + 1} 门课程：教室值为空`);
                return;
              }
              // 将内部变量重命名，避免冲突
              var currentRoomIndexInRoomsArray = rooms.indexOf(room);
              if (currentRoomIndexInRoomsArray === -1) {
                Logger.log(`跳过校区 ${campusName} 的 ${day} 第 ${rowIndex + 1} 门课程：教室 ${room} 未在 rooms 数组 [${rooms}] 中找到`);
                return;
              }
      
              var startRow = timeSlots.findIndex(t => t.getTime() >= start.getTime());
              var endRow = timeSlots.findIndex(t => t.getTime() >= end.getTime());
              // 使用重命名后的变量 currentRoomIndexInRoomsArray
              Logger.log(`校区 ${campusName} 的 ${day} 第 ${rowIndex + 1} 门课程：开始时间=${formatTime(start)} (startRow=${startRow}), 结束时间=${formatTime(end)} (endRow=${endRow}), 教室=${room} (roomIndex=${currentRoomIndexInRoomsArray})`);
              if (startRow === -1 || endRow === -1) {
                Logger.log(`跳过校区 ${campusName} 的 ${day} 第 ${rowIndex + 1} 门课程：时间槽索引无效，startRow=${startRow}, endRow=${endRow}`);
                return;
              }
      
              var currentColor = colors[colorIndex % colors.length];
              colorIndex++;
      
              var contents = [
                `[${formatTime(start)} - ${formatTime(end)}]`,
                row[headers.indexOf('Teacher')] || '',
                row[courseNameIndex] || '',
                row[headers.indexOf('Course Type')] || '',
                row[headers.indexOf('Notes')] || ''
              ];
              Logger.log(`填充校区 ${campusName} 的 ${day} 第 ${rowIndex + 1} 门课程内容：${contents}`);
      
              for (var i = startRow; i < endRow && i < timeSlots.length; i++) {
                if (i - startRow < contents.length) {
                  // 使用重命名后的变量 currentRoomIndexInRoomsArray
                  scheduleData[i][currentRoomIndexInRoomsArray] = contents[i - startRow];
                  Logger.log(`写入 scheduleData[${i}][${currentRoomIndexInRoomsArray}] = ${contents[i - startRow]}`);
                }
                // 使用重命名后的变量 currentRoomIndexInRoomsArray
                backgroundColors[i][currentRoomIndexInRoomsArray] = currentColor;
                Logger.log(`设置背景颜色：backgroundColors[${i}][${currentRoomIndexInRoomsArray}] = ${currentColor}`);
              }
            });

      // 设置表头
      var headerRow = ['时间', ...rooms];
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
      Logger.log(`校区 ${campusName} 的 ${day} sheet 表头：${headerRow}`);

      // 设置时间列和课表数据
      var timeLabels = timeSlots.map(t => formatTime(t));
      var dataRange = sheet.getRange(2, 1, timeSlots.length, headerRow.length);
      var dataValues = timeSlots.map((_, i) => [timeLabels[i], ...scheduleData[i]]);
      dataRange.setValues(dataValues);
      Logger.log(`校区 ${campusName} 的 ${day} sheet 写入 ${timeSlots.length} 行数据，列数=${headerRow.length}`);

      // 设置背景颜色
      for (var i = 0; i < timeSlots.length; i++) {
        for (var j = 0; j < rooms.length; j++) {
          if (backgroundColors[i][j]) {
            sheet.getRange(i + 2, j + 2).setBackground(backgroundColors[i][j]);
            Logger.log(`设置背景颜色：sheet[${i + 2}][${j + 2}] = ${backgroundColors[i][j]}`);
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
    });

    // 删除可能存在的临时 sheet (如果之前创建了)
    var tempSheet = spreadsheetToUse.getSheetByName("temp_placeholder");
    if (tempSheet) {
        spreadsheetToUse.deleteSheet(tempSheet);
        Logger.log("删除了临时 placeholder sheet");
    }
    // 如果是新创建的文件，删除默认的 "Sheet1" (如果它不是第一个被重命名的 sheet)
    if (!isExistingFile) {
        var defaultSheet = spreadsheetToUse.getSheetByName('Sheet1');
        // 只有当 Sheet1 存在且不是我们刚刚创建的第一个工作表 (比如 Monday) 时才删除
        if (defaultSheet && spreadsheetToUse.getSheets().length > days.length) {
            spreadsheetToUse.deleteSheet(defaultSheet);
            Logger.log("删除了新文件中的默认 Sheet1");
        }
    }
    
    // 处理文件：重命名或移动
    if (existingFile) {
      // 重命名现有文件以更新时间戳
      existingFile.setName(newFileName);
      Logger.log(`已更新文件名为：${newFileName}`);
      results.push({
        campusName,
        status: 'success',
        message: `成功更新校区 ${campusName} 的课表，文件名为 ${newFileName}`
      });
    } else {
      // 移动新创建的文件到指定文件夹
      var file = DriveApp.getFileById(spreadsheetToUse.getId());
      file.moveTo(folder);
      Logger.log(`移动文件 ${newFileName} 到文件夹 ${folderId}`);
      results.push({
        campusName,
        status: 'success',
        message: `成功为校区 ${campusName} 生成课表，文件名为 ${newFileName}`
      });
    }
    Logger.log(`处理校区 ${campusName} 课表完成`);
  } catch (e) {
    Logger.log(`为校区 ${campusName} 生成课表失败：${e.message}`);
    results.push({
      campusName,
      status: 'error',
      message: `为校区 ${campusName} 生成课表失败：${e.message}`
    });
  }
}


/**
 * 处理多个 Google Sheets 的数据，生成教师课表和校区课表
 * @param {string} folderId 输出文件夹的 Google Drive ID
 * @param {Object[]} sheetData 包含 [{ sheetLink, campusName }, ...] 的数组
 * @returns {string} 按行分隔的处理结果
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
        Logger.log(`处理表格 ${index + 1}:${sheetLink}（校区：${campusName})`);

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
        var controlData = controlSheet.getRange('A1:B3').getValues();
        if (controlData[0][0] !== 'Flag' || controlData[0][1] !== 'Value') {
          results.push({
            campusName,
            status: 'error',
            message: `校区 ${campusName} 的 "Control" 表表头无效。`
          });
          return;
        }
        var generateTeacherSchedule = controlData[1][0] === 'generateTeacherSchedule?' && controlData[1][1].toString().toLowerCase() === 'yes';
        var generateCampusSchedule = controlData[2][0] === 'generateCampusSchedule?' && controlData[2][1].toString().toLowerCase() === 'yes';

        if (!generateTeacherSchedule && !generateCampusSchedule) {
          Logger.log(`跳过 ${campusName}：generateTeacherSchedule? 和 generateCampusSchedule? 均不是 "Yes"。`);
          results.push({
            campusName,
            status: 'skipped',
            message: `因为 Control flag 均不是 Yes，跳过校区 ${campusName} 的处理。`
          });
          return;
        }

        // 生成校区课表（如果需要）
        if (generateCampusSchedule) {
          processGenerateCampusSchedule(spreadsheet, campusName, folderId, results);
        }

        // 以下为原教师课表生成逻辑
        if (!generateTeacherSchedule) {
          Logger.log(`跳过 ${campusName} 的教师课表生成：generateTeacherSchedule? 不是 "Yes"。`);
          results.push({
            campusName,
            status: 'skipped',
            message: `因为 generateTeacherSchedule? 不是 Yes，跳过校区 ${campusName} 的教师课表生成。`
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
            message: `校区 ${campusName} 的 "Teacher Data" 表缺少必要表头: Teacher, Email, NeedUpdate?`
          });
          return;
        }

        // 获取数据（从第 2 行到最后一行，仅读取需要的列）
        var lastRow = teacherSheet.getLastRow();
        var teacherData = [];
        if (lastRow > 1) {
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
              emailToTeacherMap[email] = teacherName;
            }
            teacherNeedUpdate[email] = teacherNeedUpdate[email] || needUpdate;
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
          if (!weekday || !validDays.includes(weekday)) {
            results.push({
              campusName,
              status: 'warning',
              message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行 weekday 无效：${weekday}，该课程未加入课表。`
            });
            continue;
          }
          var row = courseData[i].slice();
          if (row[startTimeIndex] instanceof Date) {
            row[startTimeIndex] = formatTime(row[startTimeIndex]);
          }
          if (row[endTimeIndex] instanceof Date) {
            row[endTimeIndex] = formatTime(row[endTimeIndex]);
          }
          var start = parseTime(row[startTimeIndex]);
          var end = parseTime(row[endTimeIndex]);
          if (!start || !end) {
            results.push({
              campusName,
              status: 'warning',
              message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行时间格式无效：开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}，该课程未加入课表。`
            });
            continue;
          }
          if (end.getTime() <= start.getTime()) {
            results.push({
              campusName,
              status: 'warning',
              message: `校区 ${campusName} 的 "Course Data" 表第 ${i + 1} 行结束时间早于或等于开始时间：开始时间=${row[startTimeIndex]}，结束时间=${row[endTimeIndex]}，该课程未加入课表。`
            });
            continue;
          }
          var durationHours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
          if (durationHours > 3) {
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
              classInfo['Campus Name'] = campusName;
              teacherSchedules[email].classes.push(classInfo);
              teacherSchedules[email].campusNames.add(campusName);
            } else if (teacher) {
              results.push({
                campusName,
                status: 'warning',
                message: `校区 ${campusName} 的 "Course Data" 表 表第 ${rowIndex + 2} 行教师 ${teacher} 未在 "Teacher Data" 表中找到，未加入课表。`
              });
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
          message: `处理校区 ${campusName} 失败：${e.message}`
        });
      }
    });

    // 检测时间冲突并标记有冲突的教师
    var teacherConflicts = {};
    Object.keys(teacherSchedules).forEach(email => {
      var { classes } = teacherSchedules[email];
      teacherConflicts[email] = { hasConflict: false, conflictDetails: [] };

      var classesByDay = {};
      classes.forEach(cls => {
        var day = cls['Weekday'];
        if (!classesByDay[day]) classesByDay[day] = [];
        classesByDay[day].push(cls);
      });

      Object.keys(classesByDay).forEach(day => {
        var dayClasses = classesByDay[day];
        dayClasses.sort((a, b) => {
          var startA = parseTime(a['Start Time']);
          var startB = parseTime(b['Start Time']);
          return startA - startB;
        });

        for (var i = 0; i < dayClasses.length; i++) {
          for (var j = i + 1; j < dayClasses.length; j++) {
            var classA = dayClasses[i];
            var classB = dayClasses[j];
            var startA = parseTime(classA['Start Time']);
            var endA = parseTime(classA['End Time']);
            var startB = parseTime(classB['Start Time']);
            var endB = parseTime(classB['End Time']);

            if (!startA || !endA || !startB || !endB) continue;

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
    var colors = ['#f6d7b0', '#b7e1cd', '#b3cde3', '#f4c7c3'];
    var colorIndex = 0;

    Object.keys(teacherSchedules).forEach(email => {
      var teacherName = emailToTeacherMap[email] || '未知教师';
      if (teacherConflicts[email].hasConflict) {
        var conflictMessages = teacherConflicts[email].conflictDetails.map(c =>
          `在 ${c.day}, ${c.classA.campus} 的课程 ${c.classA.course}(${c.classA.time}）与 ${c.classB.campus} 的课程 ${c.classB.course}(${c.classB.time}）时间冲突`
        );
        results.push({
          campusName: '教师课表',
          status: 'error',
          message: `无法为教师 ${email}(${teacherName}）生成课表：${conflictMessages.join(';')}`
        });
        return;
      }

      var { classes, campusNames } = teacherSchedules[email];
      if (!teacherNeedUpdate[email]) {
        results.push({
          campusName: '教师课表',
          status: 'skipped',
          message: `未为教师 ${email}(${teacherName}）生成课表：所有校区的 NeedUpdate? 均未设置为 Yes`
        });
        return;
      }
      if (classes.length === 0) {
        results.push({
          campusName: '教师课表',
          status: 'warning',
          message: `无法为教师 ${email}(${teacherName}）生成课表：未找到有效课程`
        });
        return;
      }

      var times = [];
      classes.forEach(cls => {
        var start = parseTime(cls['Start Time']);
        var end = parseTime(cls['End Time']);
        if (start && end) {
          times.push(start, end);
        }
      });

      if (times.length === 0) {
        results.push({
          campusName: '教师课表',
          status: 'warning',
          message: `无法为教师 ${email}(${teacherName}）生成课表：未找到有效课程时间`
        });
        return;
      }

      var minTime = new Date(Math.min(...times));
      var maxTime = new Date(Math.max(...times));
      minTime.setSeconds(0, 0);
      minTime.setMinutes(Math.floor(minTime.getMinutes() / 15) * 15);
      maxTime.setSeconds(0, 0);
      maxTime.setMinutes(Math.ceil(maxTime.getMinutes() / 15) * 15);

      var timeSlots = [];
      var current = new Date(minTime);
      while (current <= maxTime) {
        timeSlots.push(new Date(current));
        current.setMinutes(current.getMinutes() + 15);
      }

      var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
      var scheduleData = timeSlots.map(() => days.map(() => ''));
      var backgroundColors = timeSlots.map(() => days.map(() => null));

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

      var fileNamePrefix = `${email}'s Weekly Schedule`;
      var iterator = folder.getFiles();
      while (iterator.hasNext()) {
        var file = iterator.next();
        var fileName = file.getName();
        if (fileName.includes(fileNamePrefix)) {
          file.setTrashed(true);
        }
      }

      var now = new Date();
      var timeStamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      var fileName = `${email}'s Weekly Schedule (${timeStamp})`;

      var newSpreadsheet = SpreadsheetApp.create(fileName);
      var sheet = newSpreadsheet.getSheets()[0];

      var headerRow = ['Time', ...days];
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

      var timeLabels = timeSlots.map(t => formatTime(t));
      var dataRange = sheet.getRange(2, 1, timeSlots.length, headerRow.length);
      var dataValues = timeSlots.map((_, i) => [timeLabels[i], ...scheduleData[i]]);
      dataRange.setValues(dataValues);

      for (var i = 0; i < timeSlots.length; i++) {
        for (var j = 0; j < days.length; j++) {
          if (backgroundColors[i][j]) {
            sheet.getRange(i + 2, j + 2).setBackground(backgroundColors[i][j]);
          }
        }
      }

      sheet.getRange(1, 1, 1, headerRow.length).setFontWeight('bold');
      sheet.getRange(2, 1, timeSlots.length, 1).setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);
      sheet.getRange(1, 1, timeSlots.length + 1, headerRow.length).setHorizontalAlignment('center');
      for (var col = 1; col <= headerRow.length; col++) {
        sheet.setColumnWidth(col, 240);
      }

      var file = DriveApp.getFileById(newSpreadsheet.getId());
      file.moveTo(folder);

      try {
        file.addViewer(email);
        results.push({
          campusName: '教师课表',
          status: 'success',
          message: `为教师 ${email}(${teacherName}）成功生成课表`
        });
      } catch (e) {
        results.push({
          campusName: '教师课表',
          status: 'warning',
          message: `为教师 ${email}(${teacherName}）生成课表成功，但分享失败：${e.message}`
        });
      }
    });

    // 汇总结果
    var successCount = results.filter(r => r.status === 'success').length;
    var errorCount = results.filter(r => r.status === 'error').length;
    var skippedCount = results.filter(r => r.status === 'skipped').length;
    var warningCount = results.filter(r => r.status === 'warning').length;

    var outputLines = [];
    outputLines.push(`处理了 ${sheetData.length} 个校区表格：${successCount} 个成功，${errorCount} 个失败，${skippedCount} 个跳过，${warningCount} 个警告。`);
    outputLines.push('');
    outputLines.push('处理详情：');
    results.forEach(result => {
      var statusText;
      switch (result.status) {
        case 'success':
          statusText = '成功';
          break;
        case 'error':
          statusText = '错误';
          break;
        case 'skipped':
          statusText = '跳过';
          break;
        case 'warning':
          statusText = '警告';
          break;
        default:
          statusText = result.status;
      }
      outputLines.push(`- ${result.campusName} (${statusText}): ${result.message}`);
    });

    var finalOutput = outputLines.join('\n');
    Logger.log(finalOutput);
    return finalOutput;
  } catch (e) {
    Logger.log(`严重错误：${e.message}`);
    var errorOutput = `处理表格失败：${e.message}\n处理详情: \n${results.map(r => `- ${r.campusName} (${r.status}): ${r.message}`).join('\n')}`;
    return errorOutput;
  }
}
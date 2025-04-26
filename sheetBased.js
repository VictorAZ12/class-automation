/**
 * 检查在相同 Weekday 下，是否有相同 Room 的时间段重叠，动态识别列顺序
 * 如果一个课程的开始时间等于另一个课程的结束时间，不算冲突
 * @param {string} weekday 当前行的 Weekday
 * @param {string} startTime 当前行的开始时间 (hh:mm)
 * @param {string} endTime 当前行的结束时间 (hh:mm)
 * @param {string} room 当前行的 Room
 * @param {Range} range 包含标题行和数据的范围（例如，A1:E11）
 * @param {number} currentRow 当前行的索引（从 1 开始）
 * @returns {string} "Clash", "OK", 或详细错误信息
 * @customfunction
 */
function CHECK_CLASH(weekday, startTime, endTime, room, range, currentRow) {
    // 输入验证
    if (!range) return "Error: Range is undefined";
    if (!weekday) return "Error: Weekday is empty";
    if (!startTime) return "Error: Start Time is empty";
    if (!endTime) return "Error: End Time is empty";
    if (!room) return "Error: Room is empty";
    if (!currentRow) return "Error: Current row is undefined";
  
    // 获取标题行（假设为 range 的第一行）
    var headers = range[0];
    if (!headers || !Array.isArray(headers)) return "Error: No headers found in range";
  
    // 查找所需列的索引
    var weekdayCol = headers.indexOf("Weekday");
    var startTimeCol = headers.indexOf("Start Time");
    var endTimeCol = headers.indexOf("End Time");
    var roomCol = headers.indexOf("Room");
  
    // 验证所有必要列是否存在
    var missingHeaders = [];
    if (weekdayCol === -1) missingHeaders.push("Weekday");
    if (startTimeCol === -1) missingHeaders.push("Start Time");
    if (endTimeCol === -1) missingHeaders.push("End Time");
    if (roomCol === -1) missingHeaders.push("Room");
    if (missingHeaders.length > 0) {
      return "Error: Missing headers - " + missingHeaders.join(", ");
    }
  
    // 将时间字符串转换为分钟，便于比较
    function timeToMinutes(timeStr) {
      if (!timeStr || typeof timeStr !== "string") {
        return { valid: false, error: "Time is empty or not a string: " + timeStr };
      }
      var parts = timeStr.split(":");
      if (parts.length !== 2) {
        return { valid: false, error: "Invalid time format (expect hh:mm): " + timeStr };
      }
      var hours = parseInt(parts[0], 10);
      var minutes = parseInt(parts[1], 10);
      if (isNaN(hours) || isNaN(minutes) || hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
        return { valid: false, error: "Invalid time values (hh:mm): " + timeStr };
      }
      return { valid: true, minutes: hours * 60 + minutes };
    }
  
    var startResult = timeToMinutes(startTime);
    if (!startResult.valid) return "Error: " + startResult.error;
    var startMinutes = startResult.minutes;
  
    var endResult = timeToMinutes(endTime);
    if (!endResult.valid) return "Error: " + endResult.error;
    var endMinutes = endResult.minutes;
  
    if (startMinutes >= endMinutes) {
      return "Error: Start Time (" + startTime + ") is not before End Time (" + endTime + ")";
    }
  
    // 遍历数据行（从第二行开始，跳过标题行）
    for (var i = 1; i < range.length; i++) {
      // 跳过当前行（currentRow 是 Google Sheets 的行号，从 1 开始）
      if (i + 1 === currentRow) continue;
  
      var otherWeekday = range[i][weekdayCol];
      var otherStartTime = range[i][startTimeCol];
      var otherEndTime = range[i][endTimeCol];
      var otherRoom = range[i][roomCol];
  
      // 仅检查相同 Weekday 和 Room
      if (otherWeekday === weekday && otherRoom === room) {
        var otherStartResult = timeToMinutes(otherStartTime);
        var otherEndResult = timeToMinutes(otherEndTime);
  
        // 跳过无效时间
        if (!otherStartResult.valid || !otherEndResult.valid) {
          continue; // 忽略无效时间，跳到下一行
        }
        if (otherStartResult.minutes >= otherEndResult.minutes) {
          continue; // 忽略无效时间范围
        }
  
        var otherStartMinutes = otherStartResult.minutes;
        var otherEndMinutes = otherEndResult.minutes;
  
        // 检查时间重叠
        // 新逻辑：排除 start1 == end2 或 start2 == end1 的情况
        if (
          (startMinutes < otherEndMinutes && endMinutes > otherStartMinutes) ||
          (otherStartMinutes < endMinutes && otherEndMinutes > startMinutes)
        ) {
          return "Clash with Row " + (i + 1) + " (" + otherStartTime + "-" + otherEndTime + ")";
        }
      }
    }
  
    return "OK";
  }
  

  /**
 * 解析课程表并生成课程列表
 * 从 Input Sheet 读取课程表，写入 Output Sheet
 */
function parseTimetable() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 获取 Input Sheet
    var inputSheet = ss.getSheetByName("Input");
    if (!inputSheet) {
      throw new Error("Input Sheet not found");
    }
  
    // 获取数据和底色
    var lastRow = inputSheet.getLastRow();
    var lastCol = inputSheet.getLastColumn();
    var data = inputSheet.getRange(1, 1, lastRow, lastCol).getValues();
    var backgrounds = inputSheet.getRange(1, 1, lastRow, lastCol).getBackgrounds();
  
    // 获取房间名（B1:G1）
    var rooms = data[0].slice(1); // 从 B1 开始，例如 ["ROOM1", "ROOM2", ...]
  
    // 准备课程列表
    var courses = [];
    
    // 遍历每个房间列（从 B 列开始）
    for (var col = 1; col < lastCol; col++) {
      var room = rooms[col - 1]; // 房间名，例如 "ROOM1"
      var currentCourse = null;
      
      // 遍历每行（从第 2 行开始，跳过标题行）
      for (var row = 1; row < lastRow; row++) {
        var cellValue = data[row][col];
        var cellBackground = backgrounds[row][col];
        var isWhiteBackground = cellBackground === "#ffffff"; // 白色底色
  
        // 检测课程开始：非白色底色且有时间格式（hh:mm-hh:mm）
        if (!isWhiteBackground && cellValue && typeof cellValue === "string" && cellValue.match(/^\d{1,2}:\d{2}-\d{1,2}:\d{2}$/)) {
          // 提取时间
          var times = cellValue.split("-");
          var startTime = times[0].trim();
          var endTime = times[1].trim();
  
          // 检查后续行获取老师和课程名
          var teacher = row + 1 < lastRow ? data[row + 1][col] : "";
          var courseName = row + 2 < lastRow ? data[row + 2][col] : "";
  
          // 确保有老师和课程名（至少 3 行）
          if (teacher && courseName) {
            currentCourse = {
              startTime: startTime,
              endTime: endTime,
              room: room,
              teacher: teacher,
              courseName: courseName,
              background: cellBackground
            };
          }
        } else if (currentCourse) {
          // 检查课程是否结束：底色变化或变白色
          if (isWhiteBackground || cellBackground !== currentCourse.background) {
            // 课程结束，添加到列表
            courses.push([
              currentCourse.startTime,
              currentCourse.endTime,
              currentCourse.room,
              currentCourse.teacher,
              currentCourse.courseName
            ]);
            currentCourse = null;
          }
        }
      }
  
      // 如果最后一个课程未结束，添加到列表
      if (currentCourse) {
        courses.push([
          currentCourse.startTime,
          currentCourse.endTime,
          currentCourse.room,
          currentCourse.teacher,
          currentCourse.courseName
        ]);
      }
    }
  
    // 获取或创建 Output Sheet
    var outputSheet = ss.getSheetByName("Output");
    if (!outputSheet) {
      outputSheet = ss.insertSheet("Output");
    } else {
      outputSheet.clear(); // 清空现有数据
    }
  
    // 写入标题
    var headers = ["Start Time", "End Time", "Room", "Teacher", "Course Name"];
    outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
    // 写入课程数据
    if (courses.length > 0) {
      outputSheet.getRange(2, 1, courses.length, headers.length).setValues(courses);
    }
  }
  
  /**
   * 调试函数：将日志写入指定工作表
   * @param {string} message 日志消息
   */
  function logToSheet(message) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("DebugLog");
    if (!logSheet) {
      logSheet = ss.insertSheet("DebugLog");
      logSheet.getRange("A1").setValue("Timestamp");
      logSheet.getRange("B1").setValue("Message");
    }
    logSheet.appendRow([new Date(), message]);
  }
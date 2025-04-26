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
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
  
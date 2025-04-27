// @Author: Victor
// @Date: 26 Apr 2025
// @Description: This is a function designed to validate course data.
/**
 * 检查在相同 Weekday 和 Room 下，是否有时间段重叠，同时检查教师是否在同一时间段被分配到不同课程
 * 如果一个课程的开始时间等于另一个课程的结束时间，不算冲突
 * @param {string} weekday 当前行的 Weekday
 * @param {string} startTime 当前行的开始时间 (hh:mm)
 * @param {string} endTime 当前行的结束时间 (hh:mm)
 * @param {string} room 当前行的 Room
 * @param {string} teacher 当前行的 Teacher
 * @param {Range} dataRange 包含标题行和数据的范围（例如，A1:E11）
 * @param {number} currentRow 当前行的索引（从 1 开始）
 * @returns {string} "OK" 或详细错误信息（如 "Clash" 或输入无效的原因）
 * @customfunction
 */
function VALIDATE(weekday, startTime, endTime, room, teacher, dataRange, currentRow) {
  // 输入验证
  if (!dataRange) return "错误：数据范围未定义";
  if (!weekday) return "错误：Weekday 为空";
  if (!startTime) return "错误：开始时间为空";
  if (!endTime) return "错误：结束时间为空";
  if (!room) return "错误：Room 为空";
  if (!teacher) return "错误：Teacher 为空";
  if (!currentRow || !Number.isInteger(currentRow) || currentRow < 1) return "错误：当前行号无效";

  // 获取标题行（第一行）
  const headers = dataRange[0];
  if (!headers || !Array.isArray(headers)) return "错误：数据范围缺少标题行";

  // 查找所需列的索引
  const weekdayCol = headers.indexOf("Weekday");
  const startTimeCol = headers.indexOf("Start Time");
  const endTimeCol = headers.indexOf("End Time");
  const roomCol = headers.indexOf("Room");
  const teacherCol = headers.indexOf("Teacher");

  // 验证所有必要列是否存在
  const missingHeaders = [];
  if (weekdayCol === -1) missingHeaders.push("Weekday");
  if (startTimeCol === -1) missingHeaders.push("Start Time");
  if (endTimeCol === -1) missingHeaders.push("End Time");
  if (roomCol === -1) missingHeaders.push("Room");
  if (teacherCol === -1) missingHeaders.push("Teacher");
  if (missingHeaders.length > 0) {
    return "错误：缺少以下必要列 - " + missingHeaders.join(", ");
  }

  // 将时间字符串转换为分钟
  function timeToMinutes(timeStr) {
    if (!timeStr || typeof timeStr !== "string") {
      return { valid: false, error: `时间为空或非字符串: ${timeStr}` };
    }
    const timeRegex = /^([0-1]?[0-9]|2[0-3]):([0-5][0-9])$/;
    if (!timeRegex.test(timeStr)) {
      return { valid: false, error: `时间格式无效（需为 hh:mm): ${timeStr}` };
    }
    const [hours, minutes] = timeStr.split(":").map(Number);
    return { valid: true, minutes: hours * 60 + minutes };
  }

  // 验证当前行时间
  const startResult = timeToMinutes(startTime);
  if (!startResult.valid) return `错误：${startResult.error}`;
  const startMinutes = startResult.minutes;

  const endResult = timeToMinutes(endTime);
  if (!endResult.valid) return `错误：${endResult.error}`;
  const endMinutes = endResult.minutes;

  if (startMinutes >= endMinutes) {
    return `错误：开始时间 (${startTime}) 必须早于结束时间 (${endTime})`;
  }

  // 将当前行的教师拆分为数组（处理多个教师的情况）
  const currentTeachers = teacher.split(",").map(t => t.trim());

  // 遍历数据行（从第二行开始，跳过标题行）
  for (let i = 1; i < dataRange.length; i++) {
    // 跳过当前行（currentRow 从 1 开始，i 是数组索引，需加 1）
    if (i + 1 === currentRow) continue;

    // 获取其他行的数据
    const otherWeekday = dataRange[i][weekdayCol];
    const otherStartTime = dataRange[i][startTimeCol];
    const otherEndTime = dataRange[i][endTimeCol];
    const otherRoom = dataRange[i][roomCol];
    const otherTeacher = dataRange[i][teacherCol];

    // 跳过空行或不完整数据
    if (!otherWeekday || !otherStartTime || !otherEndTime || !otherRoom || !otherTeacher) {
      continue;
    }

    // 验证其他行时间
    const otherStartResult = timeToMinutes(otherStartTime);
    const otherEndResult = timeToMinutes(otherEndTime);
    if (!otherStartResult.valid || !otherEndResult.valid) {
      continue;
    }
    if (otherStartResult.minutes >= otherEndResult.minutes) {
      continue;
    }

    const otherStartMinutes = otherStartResult.minutes;
    const otherEndMinutes = otherEndResult.minutes;

    // 检查时间是否重叠（排除 start1 == end2 或 start2 == end1 的情况）
    const hasTimeOverlap = !(endMinutes <= otherStartMinutes || startMinutes >= otherEndMinutes);

    if (!hasTimeOverlap) continue; // 如果时间不重叠，直接跳过

    // 检查相同 Weekday 和 Room 的冲突（原有逻辑）
    if (otherWeekday === weekday && otherRoom === room) {
      return `冲突：与第 ${i + 1} 行（${otherWeekday}, ${otherRoom}, ${otherStartTime}-${otherEndTime}）时间重叠`;
    }

    // 检查教师冲突（即使 Room 不同）
    if (otherWeekday === weekday) {
      const otherTeachers = otherTeacher.split(",").map(t => t.trim());
      // 检查是否有共同的教师
      const commonTeachers = currentTeachers.filter(t => otherTeachers.includes(t));
      if (commonTeachers.length > 0) {
        return `教师冲突：${commonTeachers.join(", ")} 在第 ${i + 1} 行（${otherWeekday}, ${otherRoom}, ${otherStartTime}-${otherEndTime}）与当前课程时间重叠`;
      }
    }
  }

  return "OK";
}
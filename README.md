
# Uplus Schedule Process Tool 使用说明书

## 概述
Uplus Schedule Process Tool 是一个基于 Google Apps Script 开发的工具，用于从 Google Sheets 中提取课程数据，生成校区和教师的课表，并将结果保存到 Google Drive 中指定文件夹。工具支持多校区数据处理、时间冲突检测以及动态文件夹选择功能。

---

## 系统要求
- **Google 账号**：需要登录 Google 账号以访问 Google Sheets 和 Google Drive。
- **浏览器**：支持现代浏览器（如 Chrome、Firefox 等）。
- **权限**：工具需要访问用户的 Google Drive 和 Google Sheets 数据，用户需授权相关权限。

---

## 功能特点
1. **多校区支持**：可同时处理多个校区的 Google Sheets 数据。
2. **动态输入**：支持动态添加或删除 Google Sheets 链接和校区名称输入框。
3. **文件夹选择**：允许用户选择 Google Drive 中的文件夹来保存生成的课表文件。
4. **课表生成**：
   - **校区课表**：按教室和时间生成校区每日课表。
   - **教师课表**：按教师生成个人周课表，支持时间冲突检测。
5. **结果展示**：提供处理日志和直接跳转到输出文件夹的功能。
6. **加载动画**：在处理数据时显示加载动画，提升用户体验。

---

## 使用步骤

### 1. 准备 Google Sheets 数据
工具需要从 Google Sheets 中读取数据，输入的 Google Sheets 需包含以下工作表：
- **Control 表**：
  - 包含以下字段：
    - `Flag` | `Value`
    - `generateTeacherSchedule?` | `Yes` 或 `No`（是否生成教师课表）
    - `generateCampusSchedule?` | `Yes` 或 `No`（是否生成校区课表）
- **Course Data 表**：
  - 包含以下字段：
    - `Validation`, `Weekday`, `Start Time`, `End Time`, `Room`, `Teacher`, `Course Name`, `Course Type`, `Student Count`, `Students`, `Notes`
  - `Weekday` 需为 `Monday` 到 `Sunday` 之一。
  - `Start Time` 和 `End Time` 格式为 `HH:MM`（如 `09:00`）。
- **Teacher Data 表**（仅在生成教师课表时需要）：
  - 包含以下字段：
    - `Teacher`, `Email`, `NeedUpdate?`
  - `NeedUpdate?` 设置为 `Yes` 表示需要为该教师生成课表。

### 2. 访问工具
- 打开工具的部署链接（由开发者提供）。
- 登录 Google 账号并授权工具访问 Google Sheets 和 Google Drive。

### 3. 输入 Google Sheets 链接和校区名称
1. 在页面上方的输入框中：
   - **Sheet Link**：输入 Google Sheets 的链接（确保链接可访问）。
   - **Campus Name**：输入校区名称（用于命名生成的课表文件）。
2. 如需处理多个校区，点击 **Add** 按钮添加更多输入框。
3. 如需删除最后一组输入框，点击 **Delete** 按钮（至少保留一组输入框）。

### 4. 选择输出文件夹
1. 在页面中部的文件夹选择区域，浏览 Google Drive 中的文件夹。
   - 点击文件夹名称选择该文件夹（高亮显示）。
   - 点击 **Up** 按钮返回上一级文件夹。
2. 选择目标文件夹后，点击 **Confirm Selection** 确认选择。
   - 确认后会显示提示：`Folder selected: [文件夹路径]`。

### 5. 处理数据
1. 点击页面底部的 **Process** 按钮开始处理。
2. 工具会验证输入并开始处理：
   - 如果输入不完整（缺少链接、校区名称或未选择文件夹），会弹出提示。
   - 处理期间会显示加载动画。

### 6. 查看结果
1. 处理完成后，会弹出一个模态框显示处理日志：
   - 日志包括成功、失败、跳过和警告信息。
   - 示例：
     ```
     处理了 2 个校区表格：1 个成功，0 个失败，1 个跳过，0 个警告。
     处理详情：
     - CampusA (成功): 成功处理校区 CampusA 的 5 门有效课程。
     - CampusB (跳过): 因为 Control flag 均不是 Yes，跳过校区 CampusB 的处理。
     ```
2. 点击 **打开输出文件夹** 按钮可跳转到 Google Drive 中的目标文件夹查看生成的课表文件。
3. 点击 **关闭** 按钮关闭模态框。

---

## 生成的文件
- **校区课表**：
  - 文件名格式：`[校区名称]_Timetable_[时间戳].gsheet`
  - 包含多个工作表（按星期一到星期日），每个工作表按时间和教室展示课程信息。
- **教师课表**：
  - 文件名格式：`[教师邮箱]'s Weekly Schedule ([时间戳]).gsheet`
  - 包含一个工作表，按时间和星期展示教师的课程安排。

---

## 注意事项
1. **输入格式**：
   - Google Sheets 链接必须有效且用户有访问权限。
   - 时间格式需为 `HH:MM`，否则可能导致课程数据被跳过。
2. **时间冲突**：
   - 工具会检测教师课表中的时间冲突，若存在冲突，则不会生成该教师的课表，并记录错误日志。
3. **权限问题**：
   - 确保 Google Sheets 和 Google Drive 文件夹的权限设置正确，否则可能导致处理失败。
4. **数据量限制**：
   - 大量数据可能导致处理时间较长，建议分批处理。

---

## 常见问题
1. **为什么处理失败？**
   - 检查 Google Sheets 链接是否有效。
   - 确保 Google Sheets 中包含必要的表和字段。
   - 查看日志中的错误信息，可能有时间格式错误或缺少必要数据。
2. **为什么某些课程未加入课表？**
   - 检查 `Weekday` 是否为有效值（`Monday` 到 `Sunday`）。
   - 确保 `Start Time` 和 `End Time` 格式正确且结束时间晚于开始时间。
3. **如何查看生成的课表？**
   - 在结果模态框中点击 **打开输出文件夹**，即可在 Google Drive 中查看生成的文件。

---

## 联系方式
如需技术支持，请联系工具开发者或管理员。

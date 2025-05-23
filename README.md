# ExcelHelper MCP Server

ExcelHelper 是一个基于 Model Context Protocol (MCP) 的服务器，提供 Excel 文件操作功能。它允许用户通过 MCP 工具执行各种 Excel 相关的操作，如打开文件、读写单元格、进行跨表查询等。

## 功能

ExcelHelper 提供以下功能：

1. 打开 Excel 文件（包括 .xlsx 和 .csv 格式）
2. 读取单元格值
3. 写入单元格值
4. 批量写入多个单元格
5. 读取单元格公式
6. 执行跨表查询
7. 验证公式

## 配置方法

1. 确保你的系统已安装 Node.js（版本 12 或更高）。

2. 克隆此仓库到本地：

   ```
   git clone https://github.com/your-repo/ExcelHelper.git
   cd ExcelHelper
   ```

3. 安装依赖：

   ```
   npm install
   ```

4. 构建项目：

   ```
   npm run build
   ```

5. 启动服务器：

   ```
   node build/index.js
   ```

## 在 Roo Code 中启动服务

要在 Roo Code 中启动 ExcelHelper 服务，请按照以下步骤操作：

1. 打开 Roo Code 界面。

2. 在命令面板中输入以下命令来启动服务：

   ```
   node F:/MCPExcelRooCode/excel-server/build/index.js
   ```

   注意：请确保路径与你的实际项目路径相匹配。

3. 服务启动后，你会在控制台看到相关的日志信息。

## 配置 MCP 服务

要将 ExcelHelper 配置为 MCP 服务，请按照以下步骤操作：

1. 打开 Roo Code 的设置文件。通常位于：
   `C:\Users\[YourUsername]\AppData\Roaming\Code\User\globalStorage\rooveterinaryinc.roo-cline\settings\mcp_settings.json`

2. 在 `mcpServers` 对象中添加 ExcelHelper 服务的配置：

   ```json
   {
     "mcpServers": {
       "ExcelHelper": {
         "command": "node",
         "args": ["F:/MCPExcelRooCode/excel-server/build/index.js"]
       }
     }
   }
   ```

   注意：请确保路径与你的实际项目路径相匹配。

3. 保存设置文件。

4. 重启 Roo Code 或重新加载窗口以使更改生效。

现在，ExcelHelper 服务应该可以作为 MCP 服务使用了。

## 使用示例

以下是使用 ExcelHelper 的一些示例：

1. 打开 Excel 文件：

   ```javascript
   const result = await mcpClient.callTool("ExcelHelper", "open_excel", {
     filePath: "path/to/your/file.xlsx"
   });
   console.log(result.content[0].text);
   ```

2. 读取单元格值：

   ```javascript
   const result = await mcpClient.callTool("ExcelHelper", "read_cell", {
     workbookId: "1",
     sheet: "Sheet1",
     cell: "A1"
   });
   console.log(result.content[0].text);
   ```

3. 写入单元格值：

   ```javascript
   const result = await mcpClient.callTool("ExcelHelper", "write_cell", {
     workbookId: "1",
     sheet: "Sheet1",
     cell: "B2",
     value: "Hello, World!"
   });
   console.log(result.content[0].text);
   ```

4. 批量写入多个单元格：

   ```javascript
   const result = await mcpClient.callTool("ExcelHelper", "write_multiple_cells", {
     workbookId: "1",
     sheet: "Sheet1",
     cells: [
       { cell: "A1", value: "Name" },
       { cell: "B1", value: "Age" },
       { cell: "A2", value: "John" },
       { cell: "B2", value: "30" }
     ]
   });
   console.log(result.content[0].text);
   ```

注意：在使用这些功能之前，请确保你已经通过 MCP 客户端连接到 ExcelHelper 服务器。

## 贡献

欢迎提交问题和拉取请求。对于重大更改，请先开issue讨论您想要改变的内容。

## 许可

[MIT](https://choosealicense.com/licenses/mit/)
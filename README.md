# ExcelHelper MCP Server

ExcelHelper 是一个基于 Model Context Protocol (MCP) 的服务器，提供 Excel 文件操作功能。它允许用户通过 MCP 工具执行各种 Excel 相关的操作，如打开文件、读写单元格、进行跨表查询等。

## 最新更新（2025年5月27日）

今天我们对 ExcelHelper 进行了以下改进：

1. 在 `excel-server/src/index.ts` 文件中添加了 "get_all_sheets" 功能的实现。
2. 修改了 "get_all_sheets" 工具的返回格式，现在它会返回一个格式化的 sheet 列表。
3. 添加了更多的日志输出，以便更好地诊断潜在问题。
4. 测试并验证了 "open_excel" 和 "get_all_sheets" 功能。
5. 分析了 Sheet1（奖励类型和格式说明）和 Sheet2（实际配置）的内容，并提供了改进建议。

## 功能

ExcelHelper 提供以下功能：

1. 打开 Excel 文件（包括 .xlsx 和 .csv 格式）
2. 读取单元格值
3. 写入单元格值
4. 批量写入多个单元格
5. 批量读取多个单元格
6. 读取单元格公式
7. 执行跨表查询
8. 验证公式
9. 获取所有 sheet 名称

## 配置方法

1. 确保你的系统已安装 Node.js（版本 12 或更高）。

2. 克隆此仓库到本地：

   ```
   git clone 本仓库
   cd ExcelHelper
   ```

3. 安装依赖：

   ```
   npm install
   ```

4. 构建项目：

   ```
   cd excel-server
   npm run build
   ```

5. 启动服务器：

   ```
   node build/index.js
   ```

注意：在执行构建和启动操作之前，请确保你已经进入 `excel-server` 目录。

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

5. 批量读取多个单元格：

    ```javascript
    const result = await mcpClient.callTool("ExcelHelper", "read_multiple_cells", {
      workbookId: "1",
      sheet: "Sheet1",
      range: "A1:B2"
    });
    console.log(JSON.parse(result.content[0].text));
    ```

6. 获取所有 sheet 名称：

    ```javascript
    const result = await mcpClient.callTool("ExcelHelper", "get_all_sheets", {
      workbookId: "1"
    });
    console.log(result.content[0].text);
    ```

注意：在使用这些功能之前，请确保你已经通过 MCP 客户端连接到 ExcelHelper 服务器。

## 贡献

欢迎提交问题和拉取请求。对于重大更改，请先开issue讨论您想要改变的内容。

## 许可

[MIT](https://choosealicense.com/licenses/mit/)
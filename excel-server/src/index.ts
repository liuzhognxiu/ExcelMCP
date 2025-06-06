#!/usr/bin/env node

/**
 * This is an MCP server that implements Excel operations.
 * It demonstrates core MCP concepts like resources and tools by allowing:
 * - Opening Excel files
 * - Reading and writing cell data
 * - Reading formulas
 * - Performing cross-sheet queries
 * - Validating formulas
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import ExcelJS from 'exceljs';
import fs from 'fs/promises';
const { Workbook } = ExcelJS;

/**
 * Simple in-memory storage for opened workbooks.
 * In a real implementation, this might use file handles or database connections.
 */
const openWorkbooks: { [id: string]: { workbook: InstanceType<typeof Workbook>, filePath: string } } = {};

/**
 * Create an MCP server with capabilities for resources (to list/read Excel files),
 * and tools (to perform Excel operations).
 */
console.log("Creating MCP server...");
const server = new Server(
  {
    name: "ExcelHelper",
    version: "0.2.0",
  },
  {
    capabilities: {
      resources: {},
      tools: {},
    },
  }
);
console.log("MCP server created successfully.");

/**
 * Handler for listing available Excel workbooks as resources.
 * Each workbook is exposed as a resource with:
 * - An excel:// URI scheme
 * - Excel MIME type
 * - Human readable name and description
 */
server.setRequestHandler(ListResourcesRequestSchema, async () => {
  console.log("Handling ListResourcesRequest...");
  return {
    resources: Object.entries(openWorkbooks).map(([id, workbook]) => ({
      uri: `excel:///${id}`,
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      name: `Workbook ${id}`,
      description: `An opened Excel workbook`
    }))
  };
});

/**
 * Handler for reading the contents of a specific Excel workbook.
 * Takes an excel:// URI and returns the workbook information.
 */
server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
  console.log("Handling ReadResourceRequest...");
  const url = new URL(request.params.uri);
  const id = url.pathname.replace(/^\//, '');
  const workbookData = openWorkbooks[id];

  if (!workbookData) {
    throw new Error(`Workbook ${id} not found`);
  }

  return {
    contents: [{
      uri: request.params.uri,
      mimeType: "application/json",
      text: JSON.stringify({
        id,
        sheetNames: workbookData.workbook.worksheets.map((sheet: any) => sheet.name)
      })
    }]
  };
});

/**
 * Handler that lists available tools.
 * Exposes a single "create_note" tool that lets clients create new notes.
 */
server.setRequestHandler(ListToolsRequestSchema, async () => {
  console.log("Handling ListToolsRequest...");
  return {
    tools: [
      {
        name: "create_worksheet",
        description: "Create a new worksheet in an existing workbook",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheetName: {
              type: "string",
              description: "Name of the new worksheet"
            }
          },
          required: ["workbookId", "sheetName"]
        }
      },
      {
        name: "add_style",
        description: "Add or modify cell style",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheet: {
              type: "string",
              description: "Name of the sheet"
            },
            cell: {
              type: "string",
              description: "Cell address (e.g., 'A1')"
            },
            style: {
              type: "object",
              description: "Style properties",
              properties: {
                fill: {
                  type: "string",
                  description: "Fill color (e.g., '#FF0000' for red)"
                },
                font: {
                  type: "object",
                  properties: {
                    bold: { type: "boolean" },
                    color: { type: "string" }
                  }
                }
              }
            }
          },
          required: ["workbookId", "sheet", "cell", "style"]
        }
      },
      {
        name: "add_table",
        description: "Add a table with modification suggestions",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheet: {
              type: "string",
              description: "Name of the sheet"
            },
            startCell: {
              type: "string",
              description: "Start cell of the table (e.g., 'A1')"
            },
            data: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  suggestion: { type: "string" },
                  priority: { type: "string", enum: ["high", "medium", "low"] }
                }
              }
            }
          },
          required: ["workbookId", "sheet", "startCell", "data"]
        }
      },
      {
        name: "open_excel",
        description: "Open an Excel file",
        inputSchema: {
          type: "object",
          properties: {
            filePath: {
              type: "string",
              description: "Path to the Excel file"
            }
          },
          required: ["filePath"]
        }
      },
      {
        name: "get_all_sheets",
        description: "Get all sheets in a workbook",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            }
          },
          required: ["workbookId"]
        }
      },
      {
        name: "read_cell",
        description: "Read a cell value",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheet: {
              type: "string",
              description: "Name of the sheet"
            },
            cell: {
              type: "string",
              description: "Cell address (e.g., 'A1')"
            }
          },
          required: ["workbookId", "sheet", "cell"]
        }
      },
      {
        name: "read_multiple_cells",
        description: "Read values from multiple cells or rows",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheet: {
              type: "string",
              description: "Name of the sheet"
            },
            range: {
              type: "string",
              description: "Cell range (e.g., 'A1:C3') or row range (e.g., '1:3')"
            }
          },
          required: ["workbookId", "sheet", "range"]
        }
      },
      {
        name: "write_cell",
        description: "Write a value to a cell",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheet: {
              type: "string",
              description: "Name of the sheet"
            },
            cell: {
              type: "string",
              description: "Cell address (e.g., 'A1')"
            },
            value: {
              type: "string",
              description: "Value to write"
            }
          },
          required: ["workbookId", "sheet", "cell", "value"]
        }
      },
      {
        name: "write_multiple_cells",
        description: "Write values to multiple cells",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheet: {
              type: "string",
              description: "Name of the sheet"
            },
            cells: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  cell: {
                    type: "string",
                    description: "Cell address (e.g., 'A1')"
                  },
                  value: {
                    type: "string",
                    description: "Value to write"
                  }
                },
                required: ["cell", "value"]
              },
              description: "Array of cell addresses and values to write"
            }
          },
          required: ["workbookId", "sheet", "cells"]
        }
      },
      {
        name: "read_formula",
        description: "Read a cell's formula",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheet: {
              type: "string",
              description: "Name of the sheet"
            },
            cell: {
              type: "string",
              description: "Cell address (e.g., 'A1')"
            }
          },
          required: ["workbookId", "sheet", "cell"]
        }
      },
      {
        name: "cross_sheet_query",
        description: "Perform a cross-sheet query",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sourceSheet: {
              type: "string",
              description: "Name of the source sheet"
            },
            targetSheet: {
              type: "string",
              description: "Name of the target sheet"
            },
            lookupColumn: {
              type: "string",
              description: "Column to use for lookup in the source sheet"
            },
            returnColumn: {
              type: "string",
              description: "Column to return from the target sheet"
            },
            lookupValue: {
              type: "string",
              description: "Value to look up"
            }
          },
          required: ["workbookId", "sourceSheet", "targetSheet", "lookupColumn", "returnColumn", "lookupValue"]
        }
      },
      {
        name: "validate_formula",
        description: "Validate a formula",
        inputSchema: {
          type: "object",
          properties: {
            workbookId: {
              type: "string",
              description: "ID of the opened workbook"
            },
            sheet: {
              type: "string",
              description: "Name of the sheet"
            },
            cell: {
              type: "string",
              description: "Cell address (e.g., 'A1')"
            }
          },
          required: ["workbookId", "sheet", "cell"]
        }
      }
    ]
  };
});
/**
 * Handler for Excel tools.
 */
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  console.log(`Handling CallToolRequest for tool: ${request.params.name}`);
  switch (request.params.name) {
    case "create_worksheet": {
      const { workbookId, sheetName } = request.params.arguments as { workbookId: string, sheetName: string };
      const workbookData = openWorkbooks[workbookId];
      if (!workbookData) {
        throw new Error("Workbook not found");
      }
      const { workbook, filePath } = workbookData;
      try {
        const newSheet = workbook.addWorksheet(sheetName);
        await workbook.xlsx.writeFile(filePath);
        return {
          content: [{
            type: "text",
            text: `Created new worksheet '${sheetName}' in workbook ${workbookId}`
          }]
        };
      } catch (error) {
        console.error('Error creating worksheet:', error);
        throw new Error(`Failed to create worksheet: ${(error as Error).message}`);
      }
    }

    case "add_style": {
      const { workbookId, sheet, cell, style } = request.params.arguments as { workbookId: string, sheet: string, cell: string, style: { fill?: string, font?: { bold?: boolean, color?: string } } };
      const workbookData = openWorkbooks[workbookId];
      if (!workbookData) {
        throw new Error("Workbook not found");
      }
      const { workbook, filePath } = workbookData;
      const worksheet = workbook.getWorksheet(sheet);
      if (!worksheet) {
        throw new Error("Sheet not found");
      }
      const targetCell = worksheet.getCell(cell);
      if (style.fill) {
        targetCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: style.fill.replace('#', '') }
        };
      }
      if (style.font) {
        const newFont: Partial<ExcelJS.Font> = {};
        if (style.font.bold !== undefined) {
          newFont.bold = style.font.bold;
        }
        if (style.font.color) {
          newFont.color = { argb: style.font.color.replace('#', '') };
        }
        targetCell.font = {
          ...targetCell.font,
          ...newFont
        };
      }
      await workbook.xlsx.writeFile(filePath);
      return {
        content: [{
          type: "text",
          text: `Applied style to cell ${cell} in sheet ${sheet} of workbook ${workbookId}`
        }]
      };
    }

    case "add_table": {
      const { workbookId, sheet, startCell, data } = request.params.arguments as { workbookId: string, sheet: string, startCell: string, data: { suggestion: string, priority: 'high' | 'medium' | 'low' }[] };
      const workbookData = openWorkbooks[workbookId];
      if (!workbookData) {
        throw new Error("Workbook not found");
      }
      const { workbook, filePath } = workbookData;
      const worksheet = workbook.getWorksheet(sheet);
      if (!worksheet) {
        throw new Error("Sheet not found");
      }
      const startCellObj = worksheet.getCell(startCell);
      const table = worksheet.addTable({
        name: 'SuggestionTable',
        ref: startCell,
        columns: [
          { name: 'Suggestion', filterButton: true },
          { name: 'Priority', filterButton: true }
        ],
        rows: data.map(item => [item.suggestion, item.priority])
      });

      // Apply colors based on priority
      data.forEach((item, index) => {
        const rowIndex = startCellObj.row + index + 1; // +1 for header row
        const priorityCell = worksheet.getCell(rowIndex, startCellObj.col + 1);
        switch (item.priority) {
          case 'high':
            priorityCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFF0000' } // Red
            };
            break;
          case 'medium':
            priorityCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFF00' } // Yellow
            };
            break;
          case 'low':
            priorityCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FF00FF00' } // Green
            };
            break;
        }
      });

      await workbook.xlsx.writeFile(filePath);
      return {
        content: [{
          type: "text",
          text: `Added table with ${data.length} suggestions to sheet ${sheet} of workbook ${workbookId}`
        }]
      };
    }

    case "open_excel": {
      const filePath = String(request.params.arguments?.filePath);
      if (!filePath) {
        throw new Error("File path is required");
      }

      try {
        // Check if file exists and is readable
        await fs.access(filePath, fs.constants.R_OK);

        const workbook = new Workbook();
        const fileExtension = filePath.split('.').pop()?.toLowerCase();

        if (fileExtension === 'csv') {
          console.log('Reading CSV file...');
          const fileContent = await fs.readFile(filePath, 'utf8');
          console.log('CSV file content:', fileContent.substring(0, 100) + '...');
          await workbook.csv.readFile(filePath);
        } else if (fileExtension === 'xlsx' || fileExtension === 'xls') {
          console.log('Reading Excel file...');
          await workbook.xlsx.readFile(filePath);
        } else {
          throw new Error("Unsupported file format. Please use CSV or Excel files.");
        }

        const workbookId = String(Object.keys(openWorkbooks).length + 1);
        openWorkbooks[workbookId] = { workbook, filePath };

        console.log(`Successfully opened ${fileExtension.toUpperCase()} file: ${filePath}`);
        console.log('Sheets in the workbook:', workbook.worksheets.map(sheet => sheet.name));

        return {
          content: [{
            type: "text",
            text: `Opened ${fileExtension.toUpperCase()} file: ${filePath}. Workbook ID: ${workbookId}. Sheets: ${workbook.worksheets.map(sheet => sheet.name).join(', ')}`
          }]
        };
      } catch (error: unknown) {
        console.error('Error opening file:', error);
        
        // If the file couldn't be opened and it's a CSV, create a new Excel file
        if (filePath.toLowerCase().endsWith('.csv')) {
          try {
            const newWorkbook = new Workbook();
            const sheet = newWorkbook.addWorksheet('Sheet1');
            
            // Add some default content
            sheet.getCell('A1').value = 'This is a new Excel file';
            
            const newFilePath = filePath.replace('.csv', '.xlsx');
            await newWorkbook.xlsx.writeFile(newFilePath);
            
            const workbookId = String(Object.keys(openWorkbooks).length + 1);
            openWorkbooks[workbookId] = { workbook: newWorkbook, filePath: newFilePath };
            
            return {
              content: [{
                type: "text",
                text: `Created new Excel file: ${newFilePath}. Workbook ID: ${workbookId}`
              }]
            };
          } catch (newFileError) {
            console.error('Error creating new Excel file:', newFileError);
            throw new Error('Failed to create new Excel file');
          }
        }
        
        if (error instanceof Error) {
          throw new Error(`Failed to open file: ${error.message}`);
        } else {
          throw new Error('Failed to open file: Unknown error');
        }
      }
    }

    case "read_cell": {
      const { workbookId, sheet, cell } = request.params.arguments as { workbookId: string, sheet: string, cell: string };
      const workbookData = openWorkbooks[workbookId];
      if (!workbookData) {
        throw new Error("Workbook not found");
      }

      const worksheet = workbookData.workbook.getWorksheet(sheet);
      if (!worksheet) {
        throw new Error("Sheet not found");
      }

      const cellValue = worksheet.getCell(cell).value;

      return {
        content: [{
          type: "text",
          text: `Value in cell ${cell} of sheet ${sheet}: ${cellValue}`
        }]
      };
    }

    case "write_cell": {
      const { workbookId, sheet, cell, value } = request.params.arguments as { workbookId: string, sheet: string, cell: string, value: string };
      const workbookData = openWorkbooks[workbookId];
      if (!workbookData) {
        throw new Error("Workbook not found");
      }

      const { workbook, filePath } = workbookData;
      const worksheet = workbook.getWorksheet(sheet);
      if (!worksheet) {
        throw new Error("Sheet not found");
      }

      worksheet.getCell(cell).value = value;

      // Save the workbook back to the original file
      await workbook.xlsx.writeFile(filePath);

      return {
        content: [{
          type: "text",
          text: `Value ${value} written to cell ${cell} of sheet ${sheet}. Saved to ${filePath}`
        }]
      };
    }

    case "write_multiple_cells": {
      const { workbookId, sheet, cells } = request.params.arguments as { workbookId: string, sheet: string, cells: { cell: string, value: string }[] };
      const workbookData = openWorkbooks[workbookId];
      if (!workbookData) {
        throw new Error("Workbook not found");
      }

      const { workbook, filePath } = workbookData;
      const worksheet = workbook.getWorksheet(sheet);
      if (!worksheet) {
        throw new Error("Sheet not found");
      }

      cells.forEach(({ cell, value }) => {
        worksheet.getCell(cell).value = value;
      });

      // Save the workbook back to the original file
      await workbook.xlsx.writeFile(filePath);

      return {
        content: [{
          type: "text",
          text: `Successfully wrote ${cells.length} cell(s) in sheet ${sheet}. Saved to ${filePath}`
        }]
      };
    }

    case "read_multiple_cells": {
      const { workbookId, sheet, range } = request.params.arguments as { workbookId: string, sheet: string, range: string };
      const workbookData = openWorkbooks[workbookId];
      if (!workbookData) {
        throw new Error("Workbook not found");
      }

      const worksheet = workbookData.workbook.getWorksheet(sheet);
      if (!worksheet) {
        throw new Error("Sheet not found");
      }

      const [startCell, endCell] = range.split(':');
      const startCellAddress = worksheet.getCell(startCell);
      const endCellAddress = worksheet.getCell(endCell);

      const startRow = startCellAddress.row;
      const endRow = endCellAddress.row;
      const startCol = startCellAddress.col;
      const endCol = endCellAddress.col;

      if (typeof startRow !== 'number' || typeof endRow !== 'number' ||
          typeof startCol !== 'number' || typeof endCol !== 'number') {
        throw new Error("Invalid cell range");
      }

      const values = [];
      for (let row = startRow; row <= endRow; row++) {
        const rowValues = [];
        for (let col = startCol; col <= endCol; col++) {
          rowValues.push(worksheet.getCell(row, col).value);
        }
        values.push(rowValues);
      }

      return {
        content: [{
          type: "text",
          text: JSON.stringify(values)
        }]
      };
    }

    case "get_all_sheets": {
      const { workbookId } = request.params.arguments as { workbookId: string };
      const workbookData = openWorkbooks[workbookId];
      if (!workbookData) {
        throw new Error("Workbook not found");
      }

      const sheetNames = workbookData.workbook.worksheets.map(sheet => sheet.name);

      const formattedText = `Available sheets:\n${sheetNames.map((name, index) => `${index + 1}. ${name}`).join('\n')}\n\nTo select a sheet, use the "read_cell" or "read_multiple_cells" tool with the desired sheet name.`;

      return {
        content: [{
          type: "text",
          text: formattedText
        }]
      };
    }

    default: {
      throw new Error("Unknown tool");
    }
  }
});

/**
 * Start the server using stdio transport.
 * This allows the server to communicate via standard input/output streams.
 */
async function main() {
  console.log("Starting Excel Helper server...");
  const transport = new StdioServerTransport();
  try {
    await server.connect(transport);
    console.log("Server connected successfully.");
  } catch (error) {
    console.error("Error connecting server:", error);
  }
}

main().catch((error) => {
  console.error("Server error:", error);
  process.exit(1);
});


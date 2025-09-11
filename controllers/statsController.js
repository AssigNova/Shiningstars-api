const express = require("express");
const mongoose = require("mongoose");
const ExcelJS = require("exceljs");
const Post = require("../models/Post"); // adjust path
const User = require("../models/User"); // adjust path

const router = express.Router();

exports.getStats = async (req, res) => {
  try {
    // Step 1: Fetch distinct keys
    const departments = await User.distinct("department");
    const categories = await Post.distinct("category");
    const participantTypes = await Post.distinct("participantType");

    // Step 2: Aggregate counts by dept, category, participantType
    const aggregation = await Post.aggregate([
      {
        $group: {
          _id: {
            department: "$department",
            category: "$category",
            participantType: "$participantType",
          },
          count: { $sum: 1 },
        },
      },
    ]);

    // Organize data for quick lookup
    const data = {};
    for (const dept of departments) {
      data[dept] = {};
      for (const cat of categories) {
        data[dept][cat] = {};
        for (const pt of participantTypes) {
          data[dept][cat][pt] = 0;
        }
      }
    }
    aggregation.forEach(({ _id, count }) => {
      if (
        data[_id.department] &&
        data[_id.department][_id.category] &&
        data[_id.department][_id.category][_id.participantType] !== undefined
      ) {
        data[_id.department][_id.category][_id.participantType] = count;
      }
    });

    // Create workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Stats");

    // Calculate total columns: dept + categories * participantTypes
    const totalCols = 1 + categories.length * participantTypes.length;

    // Set column widths (no keys needed when setting manually)
    const columns = [];
    columns.push({ width: 25 }); // Department column wider for long names
    for (let i = 1; i < totalCols; i++) {
      columns.push({ width: 18 }); // Category/participant columns
    }
    ws.columns = columns;

    // Merge and style Department header (A1:A3)
    ws.mergeCells(1, 1, 3, 1);
    const deptHeader = ws.getCell("A1");
    deptHeader.value = "Department";
    deptHeader.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFB6C1" }, // Light pink
    };
    deptHeader.font = { bold: true, color: { argb: "FFFFFFFF" } };
    deptHeader.alignment = { vertical: "middle", horizontal: "center" };

    // Create category headers and participant type sub-headers
    let colIndex = 2;
    for (const cat of categories) {
      const startCol = colIndex;
      const endCol = colIndex + participantTypes.length - 1;

      // Merge category header row (row 1)
      ws.mergeCells(1, startCol, 1, endCol);
      const catHeader = ws.getCell(1, startCol);
      catHeader.value = cat;
      catHeader.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF4682B4" }, // Steel blue
      };
      catHeader.font = { bold: true, color: { argb: "FFFFFFFF" } };
      catHeader.alignment = { vertical: "middle", horizontal: "center" };

      // Participant type headers row 2-3 merged
      for (let i = 0; i < participantTypes.length; i++) {
        const ptCell = ws.getCell(2, colIndex + i);
        ptCell.value = participantTypes[i];
        ptCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF87CEEB" }, // Sky blue
        };
        ptCell.font = { bold: true };
        ptCell.alignment = { vertical: "middle", horizontal: "center" };
        ws.mergeCells(2, colIndex + i, 3, colIndex + i);
      }
      colIndex += participantTypes.length;
    }

    // Style department column cells (rows 4+)
    for (let r = 4; r < 4 + departments.length; r++) {
      const cell = ws.getCell(r, 1);
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF0E68C" }, // Khaki yellow
      };
      cell.alignment = { vertical: "middle", horizontal: "left" };
    }

    // Fill data starting at row 4
    let rowIndex = 4;
    for (const dept of departments) {
      ws.getCell(rowIndex, 1).value = dept;
      colIndex = 2;
      for (const cat of categories) {
        for (const pt of participantTypes) {
          ws.getCell(rowIndex, colIndex).value = data[dept][cat][pt] || 0;
          colIndex++;
        }
      }
      rowIndex++;
    }

    // Set headers to prompt file download
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="stats_export.xlsx"');

    // Write workbook to response stream & end
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Server Error");
  }
};

// exports.getStats = async (req, res) => {
//   try {
//     // Step 1: Fetch distinct data keys
//     const departments = await User.distinct("department");
//     const categories = await Post.distinct("category");
//     const participantTypes = await Post.distinct("participantType");

//     // Step 2: Aggregate post counts by department, category, participantType
//     const aggregation = await Post.aggregate([
//       {
//         $group: {
//           _id: {
//             department: "$department",
//             category: "$category",
//             participantType: "$participantType",
//           },
//           count: { $sum: 1 },
//         },
//       },
//     ]);

//     // Organize data in nested object for fast lookup
//     const data = {};
//     for (const dept of departments) {
//       data[dept] = {};
//       for (const cat of categories) {
//         data[dept][cat] = {};
//         for (const pt of participantTypes) {
//           data[dept][cat][pt] = 0;
//         }
//       }
//     }
//     aggregation.forEach(({ _id, count }) => {
//       if (
//         data[_id.department] &&
//         data[_id.department][_id.category] &&
//         data[_id.department][_id.category][_id.participantType] !== undefined
//       ) {
//         data[_id.department][_id.category][_id.participantType] = count;
//       }
//     });

//     // Step 3: Create Excel workbook and worksheet
//     const workbook = new ExcelJS.Workbook();
//     const ws = workbook.addWorksheet("Stats");

//     // Step 4: Header setup
//     // Department header cell
//     ws.mergeCells(1, 1, 3, 1); // Merge rows 1-3, column 1
//     const deptHeaderCell = ws.getCell("A1");
//     deptHeaderCell.value = "Department";
//     deptHeaderCell.fill = {
//       type: "pattern",
//       pattern: "solid",
//       fgColor: { argb: "FFB6C1" }, // Light pink
//     };
//     deptHeaderCell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white
//     deptHeaderCell.alignment = { vertical: "middle", horizontal: "center" };

//     // Category and participant type headers
//     let colIndex = 2;
//     for (const cat of categories) {
//       const startCol = colIndex;
//       const endCol = colIndex + participantTypes.length - 1;

//       // Merge category header cells
//       ws.mergeCells(1, startCol, 1, endCol);
//       const categoryCell = ws.getCell(1, startCol);
//       categoryCell.value = cat;
//       categoryCell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: "FF4682B4" }, // Steel blue
//       };
//       categoryCell.font = { bold: true, color: { argb: "FFFFFFFF" } };
//       categoryCell.alignment = { vertical: "middle", horizontal: "center" };

//       // Participant type headers (rows 2-3 merged)
//       for (let i = 0; i < participantTypes.length; i++) {
//         const ptCell = ws.getCell(2, colIndex + i);
//         ptCell.value = participantTypes[i];
//         ptCell.fill = {
//           type: "pattern",
//           pattern: "solid",
//           fgColor: { argb: "FF87CEEB" }, // Sky blue
//         };
//         ptCell.font = { bold: true };
//         ptCell.alignment = { vertical: "middle", horizontal: "center" };
//         ws.mergeCells(2, colIndex + i, 3, colIndex + i);
//       }

//       colIndex += participantTypes.length;
//     }

//     // Style department cells in column A (rows 4+)
//     for (let r = 4; r < 4 + departments.length; r++) {
//       const cell = ws.getCell(r, 1);
//       cell.fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: "FFF0E68C" }, // Khaki yellow
//       };
//       cell.alignment = { vertical: "middle", horizontal: "left" };
//     }

//     // Step 5: Fill data starting from row 4
//     let rowIndex = 4;
//     for (const dept of departments) {
//       ws.getCell(rowIndex, 1).value = dept;
//       colIndex = 2;
//       for (const cat of categories) {
//         for (const pt of participantTypes) {
//           ws.getCell(rowIndex, colIndex).value = data[dept][cat][pt] || 0;
//           colIndex++;
//         }
//       }
//       rowIndex++;
//     }

//     // Step 6: Set response headers for file download
//     res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//     res.setHeader("Content-Disposition", 'attachment; filename="stats_export.xlsx"');

//     // Step 7: Write Excel to response stream
//     await workbook.xlsx.write(res);
//     res.end();
//   } catch (err) {
//     console.error(err);
//     res.status(500).send("Server Error");
//   }
// };

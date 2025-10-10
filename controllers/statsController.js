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

    // // Step 3: Aggregate additional stats by department (posts)
    // // Total Entries, Unique Entries (count distinct titles), Total Likes, Total Comments
    // const deptStats = await Post.aggregate([
    //   {
    //     $group: {
    //       _id: "$department",
    //       totalEntries: { $sum: 1 },
    //       uniqueEntries: { $addToSet: "$title" }, // Will count size later
    //       totalLikes: { $sum: { $size: { $ifNull: ["$likes", []] } } },
    //       totalComments: { $sum: { $size: { $ifNull: ["$comments", []] } } },
    //     },
    //   },
    //   {
    //     $project: {
    //       totalEntries: 1,
    //       uniqueEntriesCount: { $size: "$uniqueEntries" },
    //       totalLikes: 1,
    //       totalComments: 1,
    //     },
    //   },
    // ]);

    // Step 3: Aggregate additional stats by department (posts)
    // Total Entries, Unique Participants (count distinct authors), Total Likes, Total Comments
    const deptStats = await Post.aggregate([
      {
        $group: {
          _id: "$department",
          totalEntries: { $sum: 1 },

          // Use author ID to get a set of unique participants
          uniqueParticipants: { $addToSet: "$author.id" },

          totalLikes: { $sum: { $size: { $ifNull: ["$likes", []] } } },
          totalComments: { $sum: { $size: { $ifNull: ["$comments", []] } } },
        },
      },
      {
        $project: {
          totalEntries: 1,

          // Rename the count to clearly indicate Unique Participants
          uniqueEntriesCount: { $size: "$uniqueParticipants" },

          totalLikes: 1,
          totalComments: 1,
        },
      },
    ]);

    // Step 4: Aggregate user count per department for total count of department
    const userCounts = await User.aggregate([
      {
        $group: {
          _id: "$department",
          totalUsers: { $sum: 1 },
        },
      },
    ]);

    // Convert userCounts to an object for quick lookup
    const userCountsMap = {};
    userCounts.forEach((userStat) => {
      userCountsMap[userStat._id] = userStat.totalUsers;
    });

    // Convert deptStats to an object for quick lookup
    const deptStatsMap = {};
    deptStats.forEach((d) => {
      deptStatsMap[d._id] = d;
    });

    // Prepare data structure for category and participant type counts
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

    // Additional columns to add at the end (based on reference excel columns requested)
    const extraColumns = [
      "Total Entries",
      "Unique Entries",
      "Total Likes",
      "Total Comments",
      "Count of Dep.",
      "Unique Participation %",
      "Participation % of Dep.",
      // "Department Scores",
    ];

    // Calculate total columns: dept + categories * participantTypes + extra columns
    const totalCols = 1 + categories.length * participantTypes.length + extraColumns.length;

    // Set column widths
    const columns = [];
    columns.push({ width: 25 }); // Department column wider
    for (let i = 1; i < 1 + categories.length * participantTypes.length; i++) {
      columns.push({ width: 18 }); // Category/participant columns
    }
    for (let i = 0; i < extraColumns.length; i++) {
      columns.push({ width: 22 }); // Extra columns widths
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

    // Create merged header cells for extra columns starting at row 1, columns after categories*participantTypes + 1
    const extraColsStart = 1 + categories.length * participantTypes.length + 1;
    ws.mergeCells(1, extraColsStart, 3, totalCols);
    const extraHeader = ws.getCell(1, extraColsStart);
    extraHeader.value = "Additional Stats";
    extraHeader.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF6A5ACD" }, // Slate Blue
    };
    extraHeader.font = { bold: true, color: { argb: "FFFFFFFF" } };
    extraHeader.alignment = { vertical: "middle", horizontal: "center" };

    // Add extra column headers on row 4
    for (let i = 0; i < extraColumns.length; i++) {
      const cell = ws.getCell(4, extraColsStart + i);
      cell.value = extraColumns[i];
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFADD8E6" }, // Light Blue
      };
      cell.font = { bold: true };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    }

    // Style department column cells (rows 5+)
    for (let r = 5; r < 5 + departments.length; r++) {
      const cell = ws.getCell(r, 1);
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF0E68C" }, // Khaki yellow
      };
      cell.alignment = { vertical: "middle", horizontal: "left" };
    }

    // Fill data starting at row 5
    let rowIndex = 5;
    for (const dept of departments) {
      ws.getCell(rowIndex, 1).value = dept;
      colIndex = 2;

      // Fill categories and participant type counts
      for (const cat of categories) {
        for (const pt of participantTypes) {
          ws.getCell(rowIndex, colIndex).value = data[dept]?.[cat]?.[pt] || 0;
          colIndex++;
        }
      }

      // Fill additional stats columns
      const stats = deptStatsMap[dept] || {
        totalEntries: 0,
        uniqueEntriesCount: 0,
        totalLikes: 0,
        totalComments: 0,
      };
      const totalUsers = userCountsMap[dept] || 0;

      ws.getCell(rowIndex, colIndex++).value = stats.totalEntries || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.uniqueEntriesCount || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.totalLikes || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.totalComments || 0;
      ws.getCell(rowIndex, colIndex++).value = totalUsers;

      // Calculate Unique Participation % by Employees if totalUsers > 0
      const uniqueParticipationPercent = totalUsers > 0 ? (stats.uniqueEntriesCount / totalUsers) * 100 : 0;
      ws.getCell(rowIndex, colIndex++).value = uniqueParticipationPercent.toFixed(2) + "%";

      // Participation % of Department - assume as uniqueEntriesCount / totalEntries to indicate participation ratio
      const participationPercent = stats.totalEntries > 0 ? (stats.uniqueEntriesCount / stats.totalEntries) * 100 : 0;
      ws.getCell(rowIndex, colIndex++).value = participationPercent.toFixed(2) + "%";

      // Department Scores - since not clearly defined, we can set as 0 or some metric, optionally you can add your logic here
      // ws.getCell(rowIndex, colIndex++).value = 0;

      rowIndex++;
    }

    // After writing all department data rows (rowIndex now points to next empty row)
    const totalsRow = ws.getRow(rowIndex);
    totalsRow.getCell(1).value = "TOTAL";
    totalsRow.getCell(1).font = { bold: true };
    totalsRow.getCell(1).alignment = { horizontal: "left" };

    // Sum columns from 2 to totalCols
    for (let col = 2; col <= totalCols; col++) {
      let sum = 0;
      // Sum from rows 5 to (rowIndex -1)
      for (let r = 5; r < rowIndex; r++) {
        const val = ws.getRow(r).getCell(col).value;
        if (typeof val === "number") {
          sum += val;
        } else if (typeof val === "string" && val.endsWith("%")) {
          // For percentage columns, ignore or handle separately if needed
        }
      }
      // Set sum cell only for numeric (skip percentage columns if preferred)
      // For simplicity, add sums only to numeric columns, skip percentage columns here
      const headerText = ws.getRow(2).getCell(col).value;
      if (headerText && !headerText.toString().includes("%")) {
        totalsRow.getCell(col).value = sum;
        totalsRow.getCell(col).font = { bold: true };
      }
    }

    totalsRow.commit();

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

exports.getStatsByParticipantType = async (req, res) => {
  try {
    // Fetch distinct keys
    const departments = await User.distinct("department");
    const categories = await Post.distinct("category");
    const participantTypes = await Post.distinct("participantType");

    // Aggregate counts by dept, category, participantType
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

    // Aggregate additional stats by department
    const deptStats = await Post.aggregate([
      {
        $group: {
          _id: "$department",
          totalEntries: { $sum: 1 },
          uniqueEntriesSet: { $addToSet: "$title" },
          totalLikes: { $sum: { $size: { $ifNull: ["$likes", []] } } },
          totalComments: { $sum: { $size: { $ifNull: ["$comments", []] } } },
        },
      },
      {
        $project: {
          totalEntries: 1,
          uniqueEntries: { $size: "$uniqueEntriesSet" },
          totalLikes: 1,
          totalComments: 1,
        },
      },
    ]);

    // Aggregate total users per department
    const userCounts = await User.aggregate([
      {
        $group: {
          _id: "$department",
          totalUsers: { $sum: 1 },
        },
      },
    ]);

    // Prepare lookup maps
    const deptStatsMap = {};
    deptStats.forEach((d) => {
      deptStatsMap[d._id] = d;
    });
    const userCountsMap = {};
    userCounts.forEach((u) => {
      userCountsMap[u._id] = u.totalUsers;
    });

    // Prepare data structure for participantType->category counts by dept
    const data = {};
    for (const dept of departments) {
      data[dept] = {};
      for (const pt of participantTypes) {
        data[dept][pt] = {};
        for (const cat of categories) {
          data[dept][pt][cat] = 0;
        }
      }
    }
    aggregation.forEach(({ _id, count }) => {
      if (
        data[_id.department] &&
        data[_id.department][_id.participantType] &&
        data[_id.department][_id.participantType][_id.category] !== undefined
      ) {
        data[_id.department][_id.participantType][_id.category] = count;
      }
    });

    // Create workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("StatsByParticipantType");

    const extraColumns = [
      "Total Entries",
      "Unique Entries",
      "Total Likes",
      "Total Comments",
      // "Total Count of Department",
      // "Unique Participation % by Employees",
      // "Participation % of Department",
    ];

    // Calculate total columns: dept + participantTypes*(categories) + extra columns
    const totalCols = 1 + participantTypes.length * categories.length + extraColumns.length;

    // Set column widths: department wider, category columns moderate, extra columns wider
    const columns = [{ width: 25 }];
    for (let i = 1; i < 1 + participantTypes.length * categories.length; i++) {
      columns.push({ width: 18 });
    }
    for (let i = 0; i < extraColumns.length; i++) {
      columns.push({ width: 22 });
    }
    ws.columns = columns;

    // Merge and style Department header (A1:A3)
    ws.mergeCells(1, 1, 3, 1);
    const deptHeader = ws.getCell("A1");
    deptHeader.value = "Department";
    deptHeader.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFB6C1" },
    };
    deptHeader.font = { bold: true, color: { argb: "FFFFFFFF" } };
    deptHeader.alignment = { vertical: "middle", horizontal: "center" };

    // Participant Type headers as top-level (Row 1)
    let colIndex = 2;
    for (const pt of participantTypes) {
      const startCol = colIndex;
      const endCol = colIndex + categories.length - 1;
      ws.mergeCells(1, startCol, 1, endCol);
      const ptHeader = ws.getCell(1, startCol);
      ptHeader.value = pt;
      ptHeader.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF4682B4" },
      };
      ptHeader.font = { bold: true, color: { argb: "FFFFFFFF" } };
      ptHeader.alignment = { vertical: "middle", horizontal: "center" };

      // Categories as sub-headers rows 2-3 merged
      for (let i = 0; i < categories.length; i++) {
        const catCell = ws.getCell(2, colIndex + i);
        catCell.value = categories[i];
        catCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF87CEEB" },
        };
        catCell.font = { bold: true };
        catCell.alignment = { vertical: "middle", horizontal: "center" };
        ws.mergeCells(2, colIndex + i, 3, colIndex + i);
      }
      colIndex += categories.length;
    }

    // Extra columns header merged and styled
    const extraColsStart = 1 + participantTypes.length * categories.length + 1;
    ws.mergeCells(1, extraColsStart, 3, totalCols);
    const extraHeader = ws.getCell(1, extraColsStart);
    extraHeader.value = "Additional Stats";
    extraHeader.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF6A5ACD" },
    };
    extraHeader.font = { bold: true, color: { argb: "FFFFFFFF" } };
    extraHeader.alignment = { vertical: "middle", horizontal: "center" };

    // Extra columns headings row 4
    for (let i = 0; i < extraColumns.length; i++) {
      const cell = ws.getCell(4, extraColsStart + i);
      cell.value = extraColumns[i];
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFADD8E6" },
      };
      cell.font = { bold: true };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    }

    // Style department cells rows 5+
    for (let r = 5; r < 5 + departments.length; r++) {
      const cell = ws.getCell(r, 1);
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF0E68C" },
      };
      cell.alignment = { vertical: "middle", horizontal: "left" };
    }

    // Write data rows starting at row 5
    let rowIndex = 5;
    for (const dept of departments) {
      ws.getCell(rowIndex, 1).value = dept;
      colIndex = 2;

      // Fill data for participantType->category
      for (const pt of participantTypes) {
        for (const cat of categories) {
          ws.getCell(rowIndex, colIndex).value = data[dept]?.[pt]?.[cat] || 0;
          colIndex++;
        }
      }

      // Fill additional stats columns
      const stats = deptStatsMap[dept] || {};
      const totalUsers = userCountsMap[dept] || 0;

      ws.getCell(rowIndex, colIndex++).value = stats.totalEntries || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.uniqueEntries || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.totalLikes || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.totalComments || 0;
      // ws.getCell(rowIndex, colIndex++).value = totalUsers;

      // const uniqueParticipationPct = totalUsers > 0 ? (stats.uniqueEntries / totalUsers) * 100 : 0;
      // ws.getCell(rowIndex, colIndex++).value = uniqueParticipationPct.toFixed(2) + "%";

      // const participationPct = stats.totalEntries > 0 ? (stats.uniqueEntries / stats.totalEntries) * 100 : 0;
      // ws.getCell(rowIndex, colIndex++).value = participationPct.toFixed(2) + "%";

      // // Department Scores (set 0 by default, customize if needed)
      // ws.getCell(rowIndex, colIndex++).value = 0;

      rowIndex++;
    }

    // Add totals row at bottom
    const totalsRow = ws.getRow(rowIndex);
    totalsRow.getCell(1).value = "TOTAL";
    totalsRow.getCell(1).font = { bold: true };
    totalsRow.getCell(1).alignment = { horizontal: "left" };

    for (let col = 2; col <= totalCols; col++) {
      let sum = 0;
      for (let r = 5; r < rowIndex; r++) {
        const val = ws.getRow(r).getCell(col).value;
        if (typeof val === "number") {
          sum += val;
        }
      }
      // Write sum for numeric columns (skip % columns)
      const headerText = ws.getRow(2).getCell(col).value;
      if (headerText && !headerText.toString().includes("%")) {
        totalsRow.getCell(col).value = sum;
        totalsRow.getCell(col).font = { bold: true };
      }
    }
    totalsRow.commit();

    // Set headers for download
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="stats_by_participantType.xlsx"');

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Server Error");
  }
};

exports.getUserStats = async (req, res) => {
  try {
    // Fetch all users for reference (ID, name, department mapping)
    const allUsers = await User.find({}, { _id: 1, name: 1, department: 1, employeeId: 1 });
    // Build a fast lookup : key is `${name}||${department}` for best matching
    function formatKey(name, dept) {
      return `${name}`.trim().toLowerCase() + "||" + `${dept}`.trim().toLowerCase();
    }

    // Build user lookup map with normalized keys
    const userLookup = {};
    allUsers.forEach((u) => {
      userLookup[formatKey(u.name, u.department)] = {
        employeeId: u.employeeId,
        name: u.name,
        department: u.department,
      };
    });
    // Aggregate likes/comments for each user found in both User/User model and any post authored
    const posts = await Post.find({});
    const userStats = {};

    // Collect all normalized keys from posts
    posts.forEach((post) => {
      const authorKey = formatKey(post.author.name, post.author.department);
      if (!userLookup[authorKey]) return; // Skip if author does not exist in User

      if (!userStats[authorKey]) {
        userStats[authorKey] = {
          employeeId: userLookup[authorKey].employeeId,
          name: post.author.name,
          department: post.author.department,
          likes: 0,
          comments: 0,
        };
      }
      userStats[authorKey].likes += Array.isArray(post.likes) ? post.likes.length : 0;
      userStats[authorKey].comments += Array.isArray(post.comments) ? post.comments.length : 0;
    });

    // Convert to array and add TOTAL
    const userArray = Object.values(userStats)
      .map((u) => ({
        ID: u.employeeId,
        Name: u.name,
        Department: u.department,
        Likes: u.likes,
        Comments: u.comments,
        TOTAL: u.likes + u.comments,
      }))
      .filter((u) => u.ID && u.Name);

    // Sort users for "top" fields
    const topByLikes = [...userArray].sort((a, b) => b.Likes - a.Likes)[0];
    const topByTotal = [...userArray].sort((a, b) => b.TOTAL - a.TOTAL)[0];

    // Generate Excel
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("User Stats");
    ws.columns = [
      { header: "ID", width: 12 },
      { header: "Name of Employee", width: 28 },
      { header: "Department", width: 23 },
      { header: "Likes", width: 14 },
      { header: "Comments", width: 14 },
      { header: "TOTAL", width: 14 },
    ];
    ws.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    ws.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF4682B4" },
    };

    // Corrected section for adding user rows
    userArray.forEach((u) => {
      ws.addRow([u.ID, u.Name, u.Department, u.Likes, u.Comments, u.TOTAL]);
    });

    ws.addRow([]);
    ws.addRow([]);

    // Footer: top users
    const footerStart = ws.lastRow.number + 1;
    ws.getCell(`C${footerStart}`).value = "BY LIKES";
    ws.getCell(`D${footerStart}`).value = topByLikes ? topByLikes.ID : "";
    ws.getCell(`E${footerStart}`).value = topByLikes ? topByLikes.Name : "";
    ws.getCell(`G${footerStart}`).value = "By Likes and comments";
    ws.getCell(`H${footerStart}`).value = topByTotal ? topByTotal.ID : "";
    ws.getCell(`I${footerStart}`).value = topByTotal ? topByTotal.Name : "";
    ws.getCell(`C${footerStart}`).font = { bold: true };
    ws.getCell(`G${footerStart}`).font = { bold: true };

    // Download
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="user_stats.xlsx"');

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Server Error");
  }
};

exports.getEntryStats = async (req, res) => {
  try {
    // Fetch all users for reference (ID, name, department mapping)
    const allUsers = await User.find({}, { _id: 1, name: 1, department: 1, employeeId: 1 });

    // Build a fast lookup: key is `${name}||${department}` for best matching
    function formatKey(name, dept) {
      return `${name}`.trim().toLowerCase() + "||" + `${dept}`.trim().toLowerCase();
    }

    // Build user lookup map with normalized keys
    const userLookup = {};
    allUsers.forEach((u) => {
      userLookup[formatKey(u.name, u.department)] = {
        employeeId: u.employeeId,
        name: u.name,
        department: u.department,
      };
    });

    // Aggregate post count for each user
    const posts = await Post.find({});
    const userStats = {};

    // Collect all normalized keys from posts and count entries
    posts.forEach((post) => {
      const authorKey = formatKey(post.author.name, post.author.department);
      if (!userLookup[authorKey]) return; // Skip if author does not exist in User

      if (!userStats[authorKey]) {
        userStats[authorKey] = {
          employeeId: userLookup[authorKey].employeeId,
          name: post.author.name,
          department: post.author.department,
          entries: 0,
        };
      }
      userStats[authorKey].entries += 1; // Increment the entry count for the author
    });

    // Convert to array
    const userArray = Object.values(userStats)
      .map((u) => ({
        ID: u.employeeId,
        Name: u.name,
        Department: u.department,
        Entries: u.entries,
      }))
      .filter((u) => u.ID && u.Name);

    // Generate Excel
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("User Entry Stats");
    ws.columns = [
      { header: "Name", width: 28 },
      { header: "Employee ID", width: 15 },
      { header: "Department", width: 23 },
      { header: "Entry", width: 14 },
    ];
    ws.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    ws.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF4682B4" },
    };

    // Add user rows
    userArray.forEach((u) => {
      ws.addRow([u.Name, u.ID, u.Department, u.Entries]);
    });

    // Download
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="entry_stats.xlsx"');

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Server Error");
  }
};

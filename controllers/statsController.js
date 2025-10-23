const express = require("express");
const mongoose = require("mongoose");
const ExcelJS = require("exceljs");
const Post = require("../models/Post"); // adjust path
const User = require("../models/User"); // adjust path

const router = express.Router();

exports.generateUserLikesReport = async (req, res) => {
  try {
    // Aggregate likes data from posts, comments, and replies
    const posts = await Post.find()
      .populate("likes", "name email employeeId department")
      .populate("comments.user", "name email employeeId department")
      .populate("comments.likes", "name email employeeId department")
      .populate("comments.replies.user", "name email employeeId department")
      .populate("comments.replies.likes", "name email employeeId department");

    // Create a map to store user like counts
    const userLikesMap = new Map();

    // Helper function to count likes
    const countLikes = (likesArray, user) => {
      if (!user) return;
      const userId = user._id.toString();
      const currentCount = userLikesMap.get(userId)?.likeCount || 0;
      userLikesMap.set(userId, {
        user: userLikesMap.get(userId)?.user || user,
        likeCount: currentCount + likesArray.length,
      });
    };

    // Process each post
    posts.forEach((post) => {
      // Count post likes
      post.likes.forEach((user) => {
        countLikes([post._id], user); // Each like is one count
      });

      // Process comments
      post.comments.forEach((comment) => {
        // Count comment likes
        comment.likes.forEach((user) => {
          countLikes([comment._id], user);
        });

        // Process replies
        comment.replies.forEach((reply) => {
          // Count reply likes
          reply.likes.forEach((user) => {
            countLikes([reply._id], user);
          });
        });
      });
    });

    // Convert map to array and sort by like count (descending)
    const userLikesData = Array.from(userLikesMap.values()).sort((a, b) => b.likeCount - a.likeCount);

    // Create Excel workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("User Likes Report");

    // Define columns
    worksheet.columns = [
      { header: "Employee ID", key: "employeeId", width: 15 },
      { header: "Name", key: "name", width: 25 },
      { header: "Email", key: "email", width: 30 },
      { header: "Department", key: "department", width: 20 },
      { header: "Total Likes Given", key: "likeCount", width: 18 },
    ];

    // Add data rows
    userLikesData.forEach((userData) => {
      worksheet.addRow({
        employeeId: userData.user.employeeId,
        name: userData.user.name,
        email: userData.user.email,
        department: userData.user.department,
        likeCount: userData.likeCount,
      });
    });

    // Style header row
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE6E6FA" },
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Set response headers for file download
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=user-likes-report.xlsx");

    // Write workbook to response
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Error generating user likes report:", error);
    res.status(500).json({
      success: false,
      message: "Failed to generate user likes report",
      error: error.message,
    });
  }
};

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

    // Step 3: Aggregate additional stats by department (posts)
    // Count UNIQUE PARTICIPANTS (users) instead of unique titles
    const deptStats = await Post.aggregate([
      {
        $group: {
          _id: "$department",
          totalEntries: { $sum: 1 },
          // Count distinct users based on author name + department
          uniqueParticipants: {
            $addToSet: {
              name: "$author.name",
              department: "$author.department",
            },
          },
          totalLikes: { $sum: { $size: { $ifNull: ["$likes", []] } } },
          totalComments: { $sum: { $size: { $ifNull: ["$comments", []] } } },
        },
      },
      {
        $project: {
          totalEntries: 1,
          uniqueParticipantsCount: { $size: "$uniqueParticipants" }, // This now counts unique users
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

    // Updated column names
    const extraColumns = [
      "Total Entries",
      "Unique Participants", // Changed from "Unique Entries"
      "Total Likes",
      "Total Comments",
      "Count of Dep.",
      "Unique Participation %",
      "Participation % of Dep.",
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
        uniqueParticipantsCount: 0, // Updated field name
        totalLikes: 0,
        totalComments: 0,
      };
      const totalUsers = userCountsMap[dept] || 0;

      ws.getCell(rowIndex, colIndex++).value = stats.totalEntries || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.uniqueParticipantsCount || 0; // Updated
      ws.getCell(rowIndex, colIndex++).value = stats.totalLikes || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.totalComments || 0;
      ws.getCell(rowIndex, colIndex++).value = totalUsers;

      // Calculate Unique Participation % by Employees if totalUsers > 0
      const uniqueParticipationPercent = totalUsers > 0 ? (stats.uniqueParticipantsCount / totalUsers) * 100 : 0;
      ws.getCell(rowIndex, colIndex++).value = uniqueParticipationPercent.toFixed(2) + "%";

      // Participation % of Department - ratio of unique participants to total entries
      const participationPercent = stats.totalEntries > 0 ? (stats.uniqueParticipantsCount / stats.totalEntries) * 100 : 0;
      ws.getCell(rowIndex, colIndex++).value = participationPercent.toFixed(2) + "%";

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

exports.getPostDetails = async (req, res) => {
  try {
    // Fetch all posts with populated author details and calculate likes/comments counts
    const posts = await Post.aggregate([
      {
        $lookup: {
          from: "users", // This should match your MongoDB collection name for users
          let: {
            authorName: "$author.name",
            authorDepartment: "$author.department",
          },
          pipeline: [
            {
              $match: {
                $expr: {
                  $and: [{ $eq: ["$name", "$$authorName"] }, { $eq: ["$department", "$$authorDepartment"] }],
                },
              },
            },
            {
              $project: {
                employeeId: 1,
                name: 1,
                department: 1,
              },
            },
          ],
          as: "userDetails",
        },
      },
      {
        $unwind: {
          path: "$userDetails",
          preserveNullAndEmptyArrays: true, // In case no matching user is found
        },
      },
      {
        $project: {
          title: 1,
          employeeId: "$userDetails.employeeId",
          authorName: "$author.name",
          department: "$author.department",
          category: 1,
          postId: "$_id",
          likesCount: { $size: { $ifNull: ["$likes", []] } },
          commentsCount: { $size: { $ifNull: ["$comments", []] } },
          createdAt: 1,
        },
      },
      {
        $sort: { createdAt: -1 }, // Sort by latest posts first
      },
    ]);

    // Alternative approach if the aggregation above doesn't work well:
    // If the aggregation is complex or slow, we can do it in two steps:
    if (posts.length === 0) {
      // Fallback: get posts and manually match with users
      const allPosts = await Post.find({}).sort({ createdAt: -1 });
      const allUsers = await User.find({}, { name: 1, department: 1, employeeId: 1 });

      // Create a lookup map for users
      const userMap = {};
      allUsers.forEach((user) => {
        const key = `${user.name.toLowerCase()}|${user.department.toLowerCase()}`;
        userMap[key] = user;
      });

      // Process posts
      const processedPosts = allPosts.map((post) => {
        const key = `${post.author.name.toLowerCase()}|${post.author.department.toLowerCase()}`;
        const user = userMap[key];

        return {
          title: post.title,
          employeeId: user ? user.employeeId : "N/A",
          authorName: post.author.name,
          department: post.author.department,
          category: post.category,
          postId: post._id,
          likesCount: Array.isArray(post.likes) ? post.likes.length : 0,
          commentsCount: Array.isArray(post.comments) ? post.comments.length : 0,
          createdAt: post.createdAt,
        };
      });

      await generateExcelReport(processedPosts, res);
    } else {
      await generateExcelReport(posts, res);
    }
  } catch (err) {
    console.error("Error in getPostDetails:", err);
    res.status(500).send("Server Error");
  }
};

// Helper function to generate Excel report
async function generateExcelReport(posts, res) {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Post Details");

  // Define columns
  ws.columns = [
    { header: "Post Title", key: "title", width: 40 },
    { header: "Employee ID", key: "employeeId", width: 15 },
    { header: "Name of User", key: "authorName", width: 25 },
    { header: "Department", key: "department", width: 20 },
    { header: "Category", key: "category", width: 20 },
    { header: "Post Link", key: "link", width: 50 },
    { header: "Total Likes", key: "likesCount", width: 12 },
    { header: "Total Comments", key: "commentsCount", width: 15 },
  ];

  // Style header row
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4682B4" }, // Steel blue
  };
  headerRow.alignment = { vertical: "middle", horizontal: "center" };

  // Add data rows
  posts.forEach((post) => {
    const postLink = `https://itcshiningstars.cosmosevents.in/posts/${post.postId}`;

    ws.addRow({
      title: post.title,
      employeeId: post.employeeId || "N/A",
      authorName: post.authorName,
      department: post.department,
      category: post.category,
      link: postLink,
      likesCount: post.likesCount,
      commentsCount: post.commentsCount,
    });
  });

  // Style the data rows
  ws.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      // Alternate row colors for better readability
      if (rowNumber % 2 === 0) {
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF0F8FF" }, // Alice blue
        };
      }

      // Center align numeric columns
      row.getCell(7).alignment = { horizontal: "center" }; // Likes
      row.getCell(8).alignment = { horizontal: "center" }; // Comments

      // Make the link cell hyperlink style
      const linkCell = row.getCell(6);
      linkCell.font = {
        color: { argb: "FF0000FF" },
        underline: true,
      };
      linkCell.value = {
        text: linkCell.value,
        hyperlink: linkCell.value,
      };
    }
  });

  // Auto-fit columns based on content
  ws.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      let columnLength = cell.value ? cell.value.toString().length : 10;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    column.width = Math.min(Math.max(maxLength + 2, column.width), 50);
  });

  // Set headers for download
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="post_details_report.xlsx"');

  await workbook.xlsx.write(res);
  res.end();
}

// // Alternative simpler version without aggregation (more reliable)
// exports.getPostDetailsSimple = async (req, res) => {
//   try {
//     // Get all posts sorted by creation date (newest first)
//     const posts = await Post.find({}).sort({ createdAt: -1 });

//     // Get all users for employeeId lookup
//     const users = await User.find({}, { name: 1, department: 1, employeeId: 1 });

//     // Create a fast lookup map for users
//     const userLookup = {};
//     users.forEach((user) => {
//       const key = `${user.name.toLowerCase()}|${user.department.toLowerCase()}`;
//       userLookup[key] = user;
//     });

//     // Process posts data
//     const postData = posts.map((post) => {
//       const userKey = `${post.author.name.toLowerCase()}|${post.author.department.toLowerCase()}`;
//       const user = userLookup[userKey];

//       return {
//         title: post.title,
//         employeeId: user ? user.employeeId : "N/A",
//         authorName: post.author.name,
//         department: post.author.department,
//         category: post.category,
//         postId: post._id,
//         likesCount: Array.isArray(post.likes) ? post.likes.length : 0,
//         commentsCount: Array.isArray(post.comments) ? post.comments.length : 0,
//         link: `https://itcshiningstars.cosmosevents.in/posts/${post._id}`,
//       };
//     });

//     // Generate Excel report
//     await generateExcelReport(postData, res);
//   } catch (err) {
//     console.error("Error in getPostDetailsSimple:", err);
//     res.status(500).send("Server Error");
//   }
// };

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

    // Aggregate additional stats by department - UPDATED for unique participants
    const deptStats = await Post.aggregate([
      {
        $group: {
          _id: "$department",
          totalEntries: { $sum: 1 },
          // Count unique participants instead of unique titles
          uniqueParticipants: {
            $addToSet: {
              name: "$author.name",
              department: "$author.department",
            },
          },
          totalLikes: { $sum: { $size: { $ifNull: ["$likes", []] } } },
          totalComments: { $sum: { $size: { $ifNull: ["$comments", []] } } },
        },
      },
      {
        $project: {
          totalEntries: 1,
          uniqueParticipantsCount: { $size: "$uniqueParticipants" }, // Updated field name
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

    // Updated column names
    const extraColumns = [
      "Total Entries",
      "Unique Participants", // Changed from "Unique Entries"
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

      // Fill additional stats columns - UPDATED for unique participants
      const stats = deptStatsMap[dept] || {
        totalEntries: 0,
        uniqueParticipantsCount: 0, // Updated field name
        totalLikes: 0,
        totalComments: 0,
      };
      const totalUsers = userCountsMap[dept] || 0;

      ws.getCell(rowIndex, colIndex++).value = stats.totalEntries || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.uniqueParticipantsCount || 0; // Updated
      ws.getCell(rowIndex, colIndex++).value = stats.totalLikes || 0;
      ws.getCell(rowIndex, colIndex++).value = stats.totalComments || 0;
      // ws.getCell(rowIndex, colIndex++).value = totalUsers;

      // If you want to include the percentage columns, uncomment and update these:
      // const uniqueParticipationPct = totalUsers > 0 ? (stats.uniqueParticipantsCount / totalUsers) * 100 : 0;
      // ws.getCell(rowIndex, colIndex++).value = uniqueParticipationPct.toFixed(2) + "%";

      // const participationPct = stats.totalEntries > 0 ? (stats.uniqueParticipantsCount / stats.totalEntries) * 100 : 0;
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

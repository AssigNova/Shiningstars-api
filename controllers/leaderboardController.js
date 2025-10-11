const Post = require("../models/Post");
const User = require("../models/User");

// Get department leaderboard
exports.getDepartmentLeaderboard = async (req, res) => {
  try {
    // Aggregate likes, submissions, participants, engagement by department
    // Only consider published posts for leaderboard calculations
    const departments = await Post.aggregate([
      { $match: { status: "published" } },
      {
        $group: {
          _id: "$department",
          submissions: { $sum: 1 },
          likes: { $sum: { $size: "$likes" } },
          participants: { $addToSet: "$author.name" },
        },
      },
      {
        $project: {
          department: "$_id",
          submissions: 1,
          likes: 1,
          participants: { $size: "$participants" },
          _id: 0,
        },
      },
      { $sort: { likes: -1 } },
    ]);
    // Calculate engagement rate as (likes / submissions) * 100, rounded
    departments.forEach((dept, i) => {
      dept.rank = i + 1;
      dept.engagement = dept.submissions > 0 ? Math.round((dept.likes / dept.submissions) * 100) : 0;
    });
    res.json(departments);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get individual leaderboard
exports.getIndividualLeaderboard = async (req, res) => {
  try {
    // Aggregate likes, submissions by user
    // Only consider published posts for individual rankings
    const individuals = await Post.aggregate([
      { $match: { status: "published" } },
      {
        $group: {
          _id: "$author.name",
          department: { $first: "$author.department" },
          submissions: { $sum: 1 },
          likes: { $sum: { $size: "$likes" } },
        },
      },
      {
        $project: {
          name: "$_id",
          department: 1,
          submissions: 1,
          likes: 1,
          _id: 0,
        },
      },
      { $sort: { likes: -1 } },
      { $limit: 5 },
    ]);
    // Add rank and badge
    individuals.forEach((person, i) => {
      person.rank = i + 1;
      person.badge = ["Top Contributor", "Rising Star", "Creative Mind", "Innovation Leader", "Community Builder"][i] || "Participant";
    });
    res.json(individuals);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get category leaders
exports.getCategoryLeaders = async (req, res) => {
  try {
    // Aggregate likes, submissions by category
    // Only consider published posts for category leaders
    const categories = await Post.aggregate([
      { $match: { status: "published" } },
      {
        $group: {
          _id: "$category",
          submissions: { $sum: 1 },
          likes: { $sum: { $size: "$likes" } },
          leader: { $first: "$author.name" },
        },
      },
      {
        $project: {
          category: "$_id",
          submissions: 1,
          likes: 1,
          leader: 1,
          _id: 0,
        },
      },
      { $sort: { likes: -1 } },
    ]);
    res.json(categories);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get submissions in the last 7 days
exports.getSubmissionsThisWeek = async (req, res) => {
  try {
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    // Count only published posts created in the last 7 days
    const count = await Post.countDocuments({
      status: "published",
      createdAt: { $gte: sevenDaysAgo },
    });
    res.json({ count });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

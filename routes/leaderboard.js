const express = require("express");
const router = express.Router();
const leaderboardController = require("../controllers/leaderboardController");

router.get("/leaderboard/departments", leaderboardController.getDepartmentLeaderboard);
router.get("/leaderboard/individuals", leaderboardController.getIndividualLeaderboard);
router.get("/leaderboard/categories", leaderboardController.getCategoryLeaders);
router.get("/leaderboard/submissionsThisWeek", leaderboardController.getSubmissionsThisWeek);

module.exports = router;

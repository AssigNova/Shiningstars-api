const express = require("express");
const router = express.Router();
const { getStats, getUserStats, getEntryStats, getStatsByParticipantType, getPostDetails } = require("../controllers/statsController");

router.get("/getStats", getStats);
router.get("/getPostsStats", getPostDetails);
router.get("/getUserStats", getUserStats);
router.get("/getEntryStats", getEntryStats);
router.get("/getStatsByParticipantType", getStatsByParticipantType);

module.exports = router;

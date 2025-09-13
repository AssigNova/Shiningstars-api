const express = require("express");
const router = express.Router();
const { getStats, getUserStats, getEntryStats, getStatsByParticipantType } = require("../controllers/statsController");

router.get("/getStats", getStats);
router.get("/getUserStats", getUserStats);
router.get("/getEntryStats", getEntryStats);
router.get("/getStatsByParticipantType", getStatsByParticipantType);

module.exports = router;

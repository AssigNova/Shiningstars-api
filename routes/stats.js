const express = require("express");
const router = express.Router();
const { getStats, getUserStats, getEntryStats } = require("../controllers/statsController");

router.get("/getStats", getStats);
router.get("/getUserStats", getUserStats);
router.get("/getEntryStats", getEntryStats);

module.exports = router;

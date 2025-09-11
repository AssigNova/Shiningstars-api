const express = require("express");
const router = express.Router();
const { getStats } = require("../controllers/statsController");

router.get("/getStats", getStats);

module.exports = router;

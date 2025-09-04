const express = require("express");
const router = express.Router();
const { requestReset, resetPassword, verifyOtp } = require("../controllers/passwordController");

router.post("/forgot", requestReset);
router.post("/verify-otp", verifyOtp);
router.post("/reset", resetPassword);

module.exports = router;

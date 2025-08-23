const express = require("express");
const router = express.Router();
const { register, login } = require("../controllers/authController");
const me = require("../controllers/meController");

router.post("/register", register);
router.post("/login", login);

// Get current user info
router.get("/me", me);

module.exports = router;

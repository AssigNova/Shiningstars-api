const express = require("express");
const router = express.Router();
const multer = require("multer");
const path = require("path");
const { createPost, getPosts, getPostsByCategory, getPostsByUser, updatePost, deletePost } = require("../controllers/postController");

// Multer storage config
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, path.join(__dirname, "../public/uploads"));
  },
  filename: function (req, file, cb) {
    const ext = path.extname(file.originalname);
    cb(null, Date.now() + "-" + file.fieldname + ext);
  },
});
const upload = multer({ storage });

// CRUD routes
router.post("/", upload.single("media"), createPost);
router.get("/", getPosts);
router.get("/category/:category", getPostsByCategory);
router.get("/user/:userId", getPostsByUser);
router.put("/:id", upload.single("media"), updatePost);
router.delete("/:id", deletePost);

// Like/Unlike/Comment routes
router.post("/:id/like", require("../controllers/postController").likePost);
router.post("/:id/unlike", require("../controllers/postController").unlikePost);
router.post("/:id/comment", require("../controllers/postController").addComment);

module.exports = router;

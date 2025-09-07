const express = require("express");
const router = express.Router();
const multer = require("multer");
const AWS = require("aws-sdk");
const { S3Client } = require("@aws-sdk/client-s3");
const multerS3 = require("multer-s3");
const path = require("path");
const {
  createPost,
  getPosts,
  getPostsByCategory,
  getPostsByUser,
  updatePost,
  deletePost,
  getPostById,
} = require("../controllers/postController");

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

// const upload = multer({ storage });

// // CRUD routes
// router.post("/", upload.single("media"), createPost);

// Create S3 instance (will use IAM Role on EC2 automatically)
const s3 = new S3Client({ region: "us-east-1" });

// const upload = multer({
//   storage,
//   limits: { fileSize: 2000 * 1024 * 1024 }, // e.g. 200MB max
//   fileFilter: (req, file, cb) => {
//     if (file.mimetype.startsWith("video/") || file.mimetype.startsWith("image/")) {
//       cb(null, true);
//     } else {
//       cb(new Error("Only video and image files are allowed"));
//     }
//   },
// });

const upload = multer({
  storage: multerS3({
    s3: s3,
    bucket: "cosmos-uploads-prod",
    // Remove the following line:
    // acl: "public-read",
    key: function (req, file, cb) {
      const fileName = Date.now().toString() + path.extname(file.originalname);
      cb(null, fileName);
    },
  }),
  limits: { fileSize: 2000 * 1024 * 1024 }, // 200 MB
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith("video/") || file.mimetype.startsWith("image/")) {
      cb(null, true);
    } else {
      cb(new Error("Only video and image files are allowed"));
    }
  },
});

router.post("/", upload.single("media"), async (req, res) => {
  try {
    await createPost(req, res);
  } catch (err) {
    res.status(500).json({ message: "Upload failed", error: err.message });
  }
});

router.get("/", getPosts);
router.get("/category/:category", getPostsByCategory);
router.get("/user/:userId", getPostsByUser);
router.put("/:id", upload.single("media"), updatePost);
router.delete("/:id", deletePost);

// Get a single post by id
router.get("/:id", getPostById);

// Views routes
const { incrementPostViews, getPostViews } = require("../controllers/postController");
router.post("/:id/view", incrementPostViews); // Increment views
router.get("/:id/views", getPostViews); // Get views

// Like/Unlike/Comment routes
router.post("/:id/like", require("../controllers/postController").likePost);
router.post("/:id/unlike", require("../controllers/postController").unlikePost);
router.post("/:id/comment", require("../controllers/postController").addComment);

// Like/unlike a comment
router.post("/:postId/comments/:commentId/like", require("../controllers/postController").likeComment);
router.post("/:postId/comments/:commentId/unlike", require("../controllers/postController").unlikeComment);

// Reply to a comment
router.post("/:postId/comments/:commentId/reply", require("../controllers/postController").replyToComment);

// Like/unlike a reply
router.post("/:postId/comments/:commentId/replies/:replyId/like", require("../controllers/postController").likeReply);
router.post("/:postId/comments/:commentId/replies/:replyId/unlike", require("../controllers/postController").unlikeReply);

module.exports = router;

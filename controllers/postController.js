const Post = require("../models/Post");

// exports.createPost = async (req, res) => {
//   try {
//     const { title, description, category, participantType, department, status, type, content, timestamp, author, likes, comments } =
//       req.body;
//     let mediaPath = null;
//     if (req.file) {
//       mediaPath = `/uploads/${req.file.filename}`;
//     }
//     // Author must be an object with name and department
//     let authorObj = author;
//     if (typeof author === "string") {
//       try {
//         authorObj = JSON.parse(author);
//       } catch {
//         authorObj = { name: author, department };
//       }
//     }

//     const post = new Post({
//       title,
//       description,
//       category,
//       author: authorObj,
//       department,
//       participantType,
//       likes: Array.isArray(likes) ? likes : [],
//       comments: Array.isArray(comments) ? comments : [],
//       timestamp,
//       status: status || "published",
//       type,
//       content: mediaPath || content,
//     });

//     await post.save();
//     res.status(201).json(post);
//   } catch (err) {
//     res.status(500).json({ message: "Server error" });
//   }
// };

exports.createPost = async (req, res) => {
  try {
    const { title, description, category, participantType, department, status, type, content, timestamp, author, likes, comments } =
      req.body;
    let mediaPath = null;
    if (req.file && req.file.location) {
      mediaPath = req.file.location; // 👈 full S3 URL
    }

    // Author must be an object with name and department
    let authorObj = author;
    if (typeof author === "string") {
      try {
        authorObj = JSON.parse(author);
      } catch {
        authorObj = { name: author, department };
      }
    }

    const post = new Post({
      title,
      description,
      category,
      author: authorObj,
      department,
      participantType,
      likes: Array.isArray(likes) ? likes : [],
      comments: Array.isArray(comments) ? comments : [],
      timestamp,
      status: status || "published",
      type,
      content: mediaPath || content,
    });

    await post.save();
    res.status(201).json(post);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get all posts
exports.getPosts = async (req, res) => {
  try {
    const posts = await Post.find()
      .populate("author", "name avatar")
      .populate("comments.user", "name department") // ✅ add this
      .sort({ createdAt: -1 });
    res.json(posts);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get posts by category
exports.getPostsByCategory = async (req, res) => {
  try {
    const { category } = req.params;
    const posts = await Post.find({ category })
      .populate("author", "name avatar")
      .populate("comments.user", "name department") // ✅ add this
      .sort({ createdAt: -1 });
    res.json(posts);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get posts by user
// Get a single post by id
exports.getPostById = async (req, res) => {
  try {
    const { id } = req.params;
    const post = await Post.findById(id).populate("author", "name avatar").populate("comments.user", "name department");
    if (!post) return res.status(404).json({ message: "Post not found" });
    res.json(post);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};
exports.getPostsByUser = async (req, res) => {
  try {
    const { userId } = req.params;
    const posts = await Post.find({ author: userId })
      .populate("author", "name avatar")
      .populate("comments.user", "name department") // ✅ add this
      .sort({ createdAt: -1 });

    res.json(posts);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Increment post views
exports.incrementPostViews = async (req, res) => {
  try {
    const { id } = req.params;
    const post = await Post.findByIdAndUpdate(id, { $inc: { views: 1 } }, { new: true });
    if (!post) return res.status(404).json({ message: "Post not found" });
    res.json({ views: post.views });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get post views
exports.getPostViews = async (req, res) => {
  try {
    const { id } = req.params;
    const post = await Post.findById(id);
    if (!post) return res.status(404).json({ message: "Post not found" });
    res.json({ views: post.views });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Update a post
exports.updatePost = async (req, res) => {
  // try {
  const { id } = req.params;
  let updateData = { ...req.body };
  // If author is a stringified object, parse it
  if (typeof updateData.author === "string") {
    try {
      updateData.author = JSON.parse(updateData.author);
    } catch {
      updateData.author = { name: updateData.author, department: updateData.department };
    }
  }
  // If file uploaded, update image/content
  if (req.file) {
    updateData.image = req.file.location;
    updateData.content = req.file.location;
  }

  const post = await Post.findByIdAndUpdate(id, updateData, { new: true });
  res.json(post);
  // } catch (err) {
  //   res.status(500).json({ message: "Server error" });
  // }
};

// Delete a post
exports.deletePost = async (req, res) => {
  try {
    const { id } = req.params;
    await Post.findByIdAndDelete(id);
    res.json({ message: "Post deleted" });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Like a post
exports.likePost = async (req, res) => {
  try {
    const { id } = req.params;
    console.log(req.body);
    const userId = req.body.userId || req.user?.userId;
    const post = await Post.findById(id);
    if (!post) return res.status(404).json({ message: "Post not found" });
    if (post.likes.includes(userId)) return res.status(400).json({ message: "Already liked" });
    post.likes.push(userId);
    await post.save();
    res.json({ message: "Post liked", likes: post.likes.length });
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: "Server error" });
  }
};

// Like a comment
exports.likeComment = async (req, res) => {
  try {
    const { postId, commentId } = req.params;
    const userId = req.body.userId || req.user?.userId;
    const post = await Post.findById(postId);
    if (!post) return res.status(404).json({ message: "Post not found" });
    const comment = post.comments.id(commentId) || post.comments.find((c) => c._id?.toString() === commentId);
    if (!comment) return res.status(404).json({ message: "Comment not found" });
    if (comment.likes.includes(userId)) return res.status(400).json({ message: "Already liked" });
    comment.likes.push(userId);
    await post.save();
    res.json({ message: "Comment liked", likes: comment.likes.length });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Unlike a comment
exports.unlikeComment = async (req, res) => {
  try {
    const { postId, commentId } = req.params;
    const userId = req.body.userId || req.user?.userId;
    const post = await Post.findById(postId);
    if (!post) return res.status(404).json({ message: "Post not found" });
    const comment = post.comments.id(commentId) || post.comments.find((c) => c._id?.toString() === commentId);
    if (!comment) return res.status(404).json({ message: "Comment not found" });
    comment.likes = comment.likes.filter((id) => id.toString() !== userId);
    await post.save();
    res.json({ message: "Comment unliked", likes: comment.likes.length });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Reply to a comment
exports.replyToComment = async (req, res) => {
  try {
    const { postId, commentId } = req.params;
    const { userId, author, department, content } = req.body;
    const post = await Post.findById(postId);
    if (!post) return res.status(404).json({ message: "Post not found" });
    const comment = post.comments.id(commentId) || post.comments.find((c) => c._id?.toString() === commentId);
    if (!comment) return res.status(404).json({ message: "Comment not found" });
    const reply = {
      user: userId,
      author,
      department,
      content,
      timestamp: new Date().toISOString(),
      likes: [],
    };
    comment.replies.push(reply);
    await post.save();
    res.json({ message: "Reply added", replies: comment.replies });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Like a reply
exports.likeReply = async (req, res) => {
  try {
    const { postId, commentId, replyId } = req.params;
    const userId = req.body.userId || req.user?.userId;
    const post = await Post.findById(postId);
    if (!post) return res.status(404).json({ message: "Post not found" });
    const comment = post.comments.id(commentId) || post.comments.find((c) => c._id?.toString() === commentId);
    if (!comment) return res.status(404).json({ message: "Comment not found" });
    const reply = comment.replies.id(replyId) || comment.replies.find((r) => r._id?.toString() === replyId);
    if (!reply) return res.status(404).json({ message: "Reply not found" });
    if (reply.likes.includes(userId)) return res.status(400).json({ message: "Already liked" });
    reply.likes.push(userId);
    await post.save();
    res.json({ message: "Reply liked", likes: reply.likes });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Unlike a reply
exports.unlikeReply = async (req, res) => {
  try {
    const { postId, commentId, replyId } = req.params;
    const userId = req.body.userId || req.user?.userId;
    const post = await Post.findById(postId);
    if (!post) return res.status(404).json({ message: "Post not found" });
    const comment = post.comments.id(commentId) || post.comments.find((c) => c._id?.toString() === commentId);
    if (!comment) return res.status(404).json({ message: "Comment not found" });
    const reply = comment.replies.id(replyId) || comment.replies.find((r) => r._id?.toString() === replyId);
    if (!reply) return res.status(404).json({ message: "Reply not found" });
    reply.likes = reply.likes.filter((id) => id.toString() !== userId);
    await post.save();
    res.json({ message: "Reply unliked", likes: reply.likes });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Unlike a post
exports.unlikePost = async (req, res) => {
  try {
    const { id } = req.params;
    const userId = req.body.userId || req.user?.userId;
    const post = await Post.findById(id);
    if (!post) return res.status(404).json({ message: "Post not found" });
    post.likes = post.likes.filter((uid) => uid.toString() !== userId);
    await post.save();
    res.json({ message: "Post unliked", likes: post.likes.length });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Add a comment to a post
exports.addComment = async (req, res) => {
  try {
    const { id } = req.params;
    const { userId, text, author, department } = req.body;
    const post = await Post.findById(id);
    if (!post) return res.status(404).json({ message: "Post not found" });
    const comment = {
      id: Date.now(),
      author,
      user: userId,
      department,
      text,
      timestamp: new Date().toLocaleString(),
      likes: [],
      replies: [],
    };
    post.comments.push(comment);
    await post.save();
    await post.populate("comments.user", "name department");
    res.json({ message: "Comment added", comments: post.comments });
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: "Server error" });
  }
};

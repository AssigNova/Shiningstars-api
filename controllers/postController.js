const Post = require("../models/Post");

// Create a new post
exports.createPost = async (req, res) => {
  try {
    const { title, description, category, participantType, department, status, type, content, timestamp, author, likes, comments } =
      req.body;
    let mediaPath = null;
    if (req.file) {
      mediaPath = `/uploads/${req.file.filename}`;
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
      likes: likes || 0,
      comments: comments || 0,
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
    const posts = await Post.find().populate("author", "name avatar").sort({ createdAt: -1 });
    res.json(posts);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get posts by category
exports.getPostsByCategory = async (req, res) => {
  try {
    const { category } = req.params;
    const posts = await Post.find({ category }).populate("author", "name avatar").sort({ createdAt: -1 });
    res.json(posts);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Get posts by user
exports.getPostsByUser = async (req, res) => {
  try {
    const { userId } = req.params;
    const posts = await Post.find({ author: userId }).populate("author", "name avatar").sort({ createdAt: -1 });
    res.json(posts);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

// Update a post
exports.updatePost = async (req, res) => {
  try {
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
      updateData.image = `/uploads/${req.file.filename}`;
      updateData.content = `/uploads/${req.file.filename}`;
    }
    const post = await Post.findByIdAndUpdate(id, updateData, { new: true });
    res.json(post);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
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
    const userId = req.body.userId || req.user?.userId;
    const post = await Post.findById(id);
    if (!post) return res.status(404).json({ message: "Post not found" });
    if (post.likes.includes(userId)) return res.status(400).json({ message: "Already liked" });
    post.likes.push(userId);
    await post.save();
    res.json({ message: "Post liked", likes: post.likes.length });
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
    const { userId, text } = req.body;
    const post = await Post.findById(id);
    if (!post) return res.status(404).json({ message: "Post not found" });
    post.comments.push({ user: userId, text });
    await post.save();
    res.json({ message: "Comment added", comments: post.comments });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

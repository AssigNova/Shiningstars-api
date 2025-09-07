const { Upload } = require("@aws-sdk/lib-storage");
const s3Client = require("../config/s3");

async function uploadToS3(file) {
  const fileName = `uploads/${Date.now()}-${file.originalname}`;

  const upload = new Upload({
    client: s3Client,
    params: {
      Bucket: "cosmos-uploads-prod",
      Key: fileName,
      Body: file.buffer,
      ContentType: file.mimetype,
    },
  });

  await upload.done();

  return `https://cosmos-uploads-prod.s3.amazonaws.com/${fileName}`;
}

module.exports = uploadToS3;

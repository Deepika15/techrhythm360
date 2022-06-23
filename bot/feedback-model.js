const mongoose = require("mongoose");

mongoose.connect("mongodb://localhost:27017/test", {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});
var db = mongoose.connection;
db.on("error", console.error.bind(console, "connection error"));
db.once("open", function (callback) {
  console.log("Connection succeeded.");
});
var Schema = mongoose.Schema;

var feedbackSchema = new Schema({
  email: String,
  collaboration: String,
  teamculture: String,
  courageous: String,
  qualityOfDilevery: String,
  commitment: String,
  recepitveness: String,
  comments: String,
});

module.exports = mongoose.model("Feedback", feedbackSchema);

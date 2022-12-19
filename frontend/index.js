const path = require("path");
const express = require("express");
const app = express();
app.use(express.static(path.join(__dirname, "public")));
app.get("/", (req, res) => {
  alert("Hello World!");
  console.log("query");
});
app.listen(3333, () => {
  console.log("Application listening on port 3333!");
});

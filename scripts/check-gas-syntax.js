const fs = require("fs");
const path = require("path");
const vm = require("vm");

const gasPath = path.join(__dirname, "..", "gas", "Code.gs");
const source = fs.readFileSync(gasPath, "utf8");

new vm.Script(source, {
  filename: gasPath,
});

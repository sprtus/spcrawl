const path = require('path');

module.exports = {
  // Build path
  path: function(filePath, cwd = true) {
    return path.normalize(`${cwd ? process.cwd() : __dirname}/${filePath}`);
  },
};

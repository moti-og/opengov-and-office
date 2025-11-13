const mongoose = require('mongoose');

const budgetBookSchema = new mongoose.Schema({
  // Legacy single image support
  image: { type: String, default: '' },
  
  // New multi-screenshot support
  screenshots: [{
    address: { type: String },  // e.g. "A1:F10"
    image: { type: String }      // base64 image string
  }],
  
  updatedAt: { type: Date, default: Date.now }
});

budgetBookSchema.pre('save', function(next) {
  this.updatedAt = Date.now();
  next();
});

module.exports = mongoose.model('BudgetBook', budgetBookSchema);


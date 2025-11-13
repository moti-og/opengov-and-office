const mongoose = require('mongoose');

const budgetBookSchema = new mongoose.Schema({
  image: { type: String, default: '' }, // base64 image string
  updatedAt: { type: Date, default: Date.now }
});

budgetBookSchema.pre('save', function(next) {
  this.updatedAt = Date.now();
  next();
});

module.exports = mongoose.model('BudgetBook', budgetBookSchema);


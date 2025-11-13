const mongoose = require('mongoose');

const budgetBookSchema = new mongoose.Schema({
  data: { type: [[String]], default: [] },
  updatedAt: { type: Date, default: Date.now }
});

budgetBookSchema.pre('save', function(next) {
  this.updatedAt = Date.now();
  next();
});

module.exports = mongoose.model('BudgetBook', budgetBookSchema);


const mongoose = require('mongoose');

const documentSchema = new mongoose.Schema({
  documentId: { type: String, required: true, unique: true },
  title: { type: String, required: true },
  type: { type: String, enum: ['excel', 'word', 'powerpoint', 'web'], required: true },
  data: { type: [[String]], default: [] },
  layout: {
    columnWidths: { type: [Number], default: [] },  // Width in pixels for each column
    rowHeights: { type: [Number], default: [] }     // Height in pixels for each row
  },
  metadata: {
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now },
    version: { type: Number, default: 1 }
  }
});

documentSchema.pre('save', function(next) {
  this.metadata.updatedAt = Date.now();
  next();
});

module.exports = mongoose.model('Document', documentSchema);


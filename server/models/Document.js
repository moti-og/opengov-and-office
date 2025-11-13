const mongoose = require('mongoose');

const documentSchema = new mongoose.Schema({
  documentId: { type: String, required: true, unique: true },
  title: { type: String, required: true },
  type: { type: String, enum: ['excel', 'word', 'powerpoint', 'web'], required: true },
  
  // Support both old (data) and new (ranges) format
  data: { type: [[String]], default: [] },  // Legacy support
  ranges: [{
    address: { type: String },  // e.g. "A1:F10"
    data: { type: [[String]] }   // 2D array
  }],
  
  layout: {
    columnWidths: { type: [Number], default: [] },  // Width in pixels for each column
    rowHeights: { type: [Number], default: [] }     // Height in pixels for each row
  },
  charts: [{
    name: { type: String },
    image: { type: String }  // base64 encoded image
  }],
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


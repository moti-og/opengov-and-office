const mongoose = require('mongoose');

const sessionSchema = new mongoose.Schema({
  sessionId: { type: String, required: true, unique: true },
  platform: { type: String, enum: ['excel', 'word', 'powerpoint', 'web'], required: true },
  documentId: { type: String, required: true },
  connectedAt: { type: Date, default: Date.now },
  status: { type: String, enum: ['active', 'disconnected'], default: 'active' }
});

module.exports = mongoose.model('Session', sessionSchema);

require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Serve static files for add-ins
app.use('/addin', express.static(path.join(__dirname, 'addin')));
app.use('/assets', express.static(path.join(__dirname, 'assets')));

mongoose.connect(process.env.MONGODB_URI)
  .then(() => console.log('Connected to MongoDB'))
  .catch(err => console.error('MongoDB connection error:', err));

const clients = [];

app.get('/', (req, res) => {
  res.json({ message: 'OpenGov Office Server Running', status: 'ok' });
});

app.get('/api/stream', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  clients.push(res);
  console.log(`Client connected. Total: ${clients.length}`);
  req.on('close', () => {
    const index = clients.indexOf(res);
    if (index !== -1) clients.splice(index, 1);
    console.log(`Client disconnected. Total: ${clients.length}`);
  });
  res.write(`data: ${JSON.stringify({ type: 'connected' })}\n\n`);
});

function broadcast(event, data) {
  const message = `data: ${JSON.stringify({ type: event, ...data })}\n\n`;
  clients.forEach(client => {
    try {
      client.write(message);
    } catch (error) {
      console.error('Broadcast error:', error);
    }
  });
  console.log(`Broadcasted ${event} to ${clients.length} clients`);
}

app.set('broadcast', broadcast);

const documentsRouter = require('./routes/documents');
app.use('/api/documents', documentsRouter);

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log(`Add-in files: http://localhost:${PORT}/addin/excel/taskpane.html`);
});

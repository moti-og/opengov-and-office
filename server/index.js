require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

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
  req.on('close', () => {
    const index = clients.indexOf(res);
    if (index !== -1) clients.splice(index, 1);
  });
  res.write(`data: ${JSON.stringify({ type: 'connected' })}\n\n`);
});

function broadcast(event, data) {
  const message = `data: ${JSON.stringify({ type: event, ...data })}\n\n`;
  clients.forEach(client => client.write(message));
}

app.set('broadcast', broadcast);

const documentsRouter = require('./routes/documents');
app.use('/api/documents', documentsRouter);

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});

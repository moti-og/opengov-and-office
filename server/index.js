const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');

const app = express();
const PORT = process.env.SERVER_PORT || 3001;

app.use(cors());
app.use(express.json());

// Serve manifest files (must be before static middleware)
app.get('/manifest.xml', (req, res) => {
  const manifestPath = path.join(__dirname, '..', 'manifest.xml');
  console.log('Manifest requested. Path:', manifestPath);
  console.log('File exists:', require('fs').existsSync(manifestPath));
  res.sendFile(manifestPath);
});
app.get('/manifest-local.xml', (req, res) => {
  const manifestPath = path.join(__dirname, '..', 'manifest-local.xml');
  console.log('Local manifest requested. Path:', manifestPath);
  console.log('File exists:', require('fs').existsSync(manifestPath));
  res.sendFile(manifestPath);
});

// Serve installer files from tools/scripts
const scriptsPath = path.join(__dirname, '..', 'tools', 'scripts');
app.get('/install-excel-addin.bat', (req, res) => res.sendFile(path.join(scriptsPath, 'install-excel-addin.bat')));
app.get('/install-excel-addin.sh', (req, res) => res.sendFile(path.join(scriptsPath, 'install-excel-addin.sh')));
app.get('/install-excel-addin-local.bat', (req, res) => res.sendFile(path.join(scriptsPath, 'install-excel-addin-local.bat')));
app.get('/install-excel-addin-local.sh', (req, res) => res.sendFile(path.join(scriptsPath, 'install-excel-addin-local.sh')));
app.get('/uninstall-excel-addin.bat', (req, res) => res.sendFile(path.join(scriptsPath, 'uninstall-excel-addin.bat')));
app.get('/uninstall-excel-addin.sh', (req, res) => res.sendFile(path.join(scriptsPath, 'uninstall-excel-addin.sh')));

// Serve addin folder
app.use('/addin', express.static(path.join(__dirname, '..', 'addin')));

// Serve web interface at root (must be last to not override other routes)
app.use(express.static(path.join(__dirname, '..', 'web')));

mongoose.connect(process.env.MONGODB_URI || 'mongodb://localhost:27017/opengov-office')
  .then(() => console.log('✓ Connected to MongoDB'))
  .catch(err => console.error('MongoDB connection error:', err));

const clients = [];

app.get('/api/health', (req, res) => {
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
  console.log(`✓ Server running on http://localhost:${PORT}`);
  console.log(`✓ SSE endpoint: http://localhost:${PORT}/api/stream`);
});


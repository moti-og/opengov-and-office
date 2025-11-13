const express = require('express');
const router = express.Router();
const Document = require('../models/Document');

router.get('/', async (req, res) => {
  try {
    const documents = await Document.find();
    res.json(documents);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

router.get('/:id', async (req, res) => {
  try {
    const document = await Document.findOne({ documentId: req.params.id });
    if (!document) return res.status(404).json({ error: 'Document not found' });
    res.json(document);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

router.post('/', async (req, res) => {
  try {
    const document = new Document(req.body);
    await document.save();
    const broadcast = req.app.get('broadcast');
    broadcast('document-created', { documentId: document.documentId });
    res.status(201).json(document);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

router.post('/:id/update', async (req, res) => {
  try {
    let document = await Document.findOne({ documentId: req.params.id });
    if (!document) {
      // Create if doesn't exist
      document = new Document({
        documentId: req.params.id,
        title: req.body.title || 'Untitled',
        type: req.body.type || 'excel',
        data: req.body.data,
        layout: req.body.layout || { columnWidths: [], rowHeights: [] },
        charts: req.body.charts || []
      });
    } else {
      // Update data if provided
      if (req.body.data !== undefined) {
        document.data = req.body.data;
      }
      
      // Update layout if provided
      if (req.body.layout !== undefined) {
        document.layout = req.body.layout;
      }
      
      // Update charts if provided
      if (req.body.charts !== undefined) {
        document.charts = req.body.charts;
      }
      
      document.metadata.version += 1;
    }
    await document.save();
    const broadcast = req.app.get('broadcast');
    broadcast('data-update', { 
      documentId: document.documentId, 
      data: document.data,
      layout: document.layout,
      charts: document.charts,
      sourceType: req.body.type || 'unknown'  // Track where update came from
    });
    res.json(document);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

module.exports = router;


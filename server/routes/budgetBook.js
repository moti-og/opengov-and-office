const express = require('express');
const router = express.Router();
const BudgetBook = require('../models/BudgetBook');

// GET budget book data
router.get('/', async (req, res) => {
  try {
    let budgetBook = await BudgetBook.findOne();
    
    if (!budgetBook) {
      return res.json({ screenshots: [], image: null, updatedAt: null });
    }
    
    // Return both new and legacy format
    res.json({
      screenshots: budgetBook.screenshots || [],
      image: budgetBook.image || null,  // Legacy support
      updatedAt: budgetBook.updatedAt
    });
  } catch (error) {
    console.error('Error fetching budget book:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST/UPDATE budget book
router.post('/update', async (req, res) => {
  try {
    const { image, screenshots } = req.body;
    
    // Validate input - need either image (legacy) or screenshots (new)
    if (!image && (!screenshots || !Array.isArray(screenshots))) {
      return res.status(400).json({ error: 'Invalid request format' });
    }
    
    let budgetBook = await BudgetBook.findOne();
    
    if (!budgetBook) {
      budgetBook = new BudgetBook({
        image: image || '',
        screenshots: screenshots || []
      });
    } else {
      if (image) {
        budgetBook.image = image;
      }
      if (screenshots) {
        budgetBook.screenshots = screenshots;
      }
    }
    
    await budgetBook.save();
    
    const count = screenshots ? screenshots.length : 1;
    console.log('Budget book updated:', count, 'screenshot(s)');
    
    res.json({
      success: true,
      updatedAt: budgetBook.updatedAt
    });
  } catch (error) {
    console.error('Error updating budget book:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;


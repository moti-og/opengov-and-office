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
    console.log('Budget book update request:', {
      hasImage: !!req.body.image,
      hasScreenshots: !!req.body.screenshots,
      screenshotsCount: req.body.screenshots?.length
    });
    
    const { image, screenshots } = req.body;
    
    // Validate input - need either image (legacy) or screenshots (new)
    if (!image && (!screenshots || !Array.isArray(screenshots))) {
      console.error('Invalid budget book request format');
      return res.status(400).json({ error: 'Invalid request format - need image or screenshots array' });
    }
    
    let budgetBook = await BudgetBook.findOne();
    
    if (!budgetBook) {
      console.log('Creating new budget book');
      budgetBook = new BudgetBook({
        image: image || '',
        screenshots: screenshots || []
      });
    } else {
      console.log('Updating existing budget book');
      if (image) {
        budgetBook.image = image;
      }
      if (screenshots) {
        budgetBook.screenshots = screenshots;
      }
    }
    
    await budgetBook.save();
    
    const count = screenshots ? screenshots.length : 1;
    console.log('Budget book updated successfully:', count, 'screenshot(s)');
    
    res.json({
      success: true,
      updatedAt: budgetBook.updatedAt
    });
  } catch (error) {
    console.error('Error updating budget book:', error);
    res.status(500).json({ error: error.message, stack: error.stack });
  }
});

module.exports = router;


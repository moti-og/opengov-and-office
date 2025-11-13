const express = require('express');
const router = express.Router();
const BudgetBook = require('../models/BudgetBook');

// GET budget book image
router.get('/', async (req, res) => {
  try {
    // Try to find the budget book data
    let budgetBook = await BudgetBook.findOne();
    
    if (!budgetBook || !budgetBook.image) {
      // Return empty if not found
      return res.json({ image: null, updatedAt: null });
    }
    
    res.json({
      image: budgetBook.image,
      updatedAt: budgetBook.updatedAt
    });
  } catch (error) {
    console.error('Error fetching budget book:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST/UPDATE budget book image
router.post('/update', async (req, res) => {
  try {
    const { image } = req.body;
    
    if (!image || typeof image !== 'string') {
      return res.status(400).json({ error: 'Invalid image format' });
    }
    
    // Find existing or create new
    let budgetBook = await BudgetBook.findOne();
    
    if (!budgetBook) {
      budgetBook = new BudgetBook({ image });
    } else {
      budgetBook.image = image;
    }
    
    await budgetBook.save();
    
    console.log('Budget book updated, image size:', image.length);
    
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


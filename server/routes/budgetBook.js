const express = require('express');
const router = express.Router();
const BudgetBook = require('../models/BudgetBook');

// GET budget book data
router.get('/', async (req, res) => {
  try {
    // Try to find the budget book data
    let budgetBook = await BudgetBook.findOne();
    
    if (!budgetBook) {
      // Return empty if not found
      return res.json({ data: [], updatedAt: null });
    }
    
    res.json({
      data: budgetBook.data,
      updatedAt: budgetBook.updatedAt
    });
  } catch (error) {
    console.error('Error fetching budget book:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST/UPDATE budget book data
router.post('/update', async (req, res) => {
  try {
    const { data } = req.body;
    
    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ error: 'Invalid data format' });
    }
    
    // Find existing or create new
    let budgetBook = await BudgetBook.findOne();
    
    if (!budgetBook) {
      budgetBook = new BudgetBook({ data });
    } else {
      budgetBook.data = data;
    }
    
    await budgetBook.save();
    
    console.log('Budget book updated:', data.length, 'rows');
    
    res.json({
      success: true,
      data: budgetBook.data,
      updatedAt: budgetBook.updatedAt
    });
  } catch (error) {
    console.error('Error updating budget book:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;


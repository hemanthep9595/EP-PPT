const express = require('express');
const router = express.Router();
const CampaignBrief = require('../models/CampaignBrief');

// POST /api/campaign-brief - Save form submission
router.post('/campaign-brief', async (req, res) => {
  try {
    const campaignBrief = new CampaignBrief(req.body);
    const savedBrief = await campaignBrief.save();

    res.status(201).json({
      success: true,
      message: 'Campaign brief submitted successfully!',
      data: {
        id: savedBrief._id,
        brandName: savedBrief.brandName,
        createdAt: savedBrief.createdAt
      }
    });
  } catch (error) {
    if (error.name === 'ValidationError') {
      const messages = Object.values(error.errors).map(err => err.message);
      return res.status(400).json({
        success: false,
        message: 'Validation failed',
        errors: messages
      });
    }

    console.error('Error saving campaign brief:', error);
    res.status(500).json({
      success: false,
      message: 'An error occurred while submitting your campaign brief. Please try again.'
    });
  }
});

// GET /api/campaign-brief/:id - Retrieve submission
router.get('/campaign-brief/:id', async (req, res) => {
  try {
    const campaignBrief = await CampaignBrief.findById(req.params.id);

    if (!campaignBrief) {
      return res.status(404).json({
        success: false,
        message: 'Campaign brief not found'
      });
    }

    res.json({
      success: true,
      data: campaignBrief
    });
  } catch (error) {
    if (error.name === 'CastError') {
      return res.status(400).json({
        success: false,
        message: 'Invalid campaign brief ID'
      });
    }

    console.error('Error retrieving campaign brief:', error);
    res.status(500).json({
      success: false,
      message: 'An error occurred while retrieving the campaign brief.'
    });
  }
});

// GET /api/campaign-briefs - List all submissions (for admin use later)
router.get('/campaign-briefs', async (req, res) => {
  try {
    const briefs = await CampaignBrief.find()
      .select('brandName industry campaignObjective status createdAt')
      .sort({ createdAt: -1 })
      .limit(100);

    res.json({
      success: true,
      count: briefs.length,
      data: briefs
    });
  } catch (error) {
    console.error('Error listing campaign briefs:', error);
    res.status(500).json({
      success: false,
      message: 'An error occurred while retrieving campaign briefs.'
    });
  }
});

module.exports = router;

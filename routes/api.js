const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const CampaignBrief = require('../models/CampaignBrief');
const { generateCustomPptx, formatCampaignBriefForPptx } = require('../utils/pptxModifier');

// Configure multer for logo uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, '..', 'uploads', 'logos');
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 }, // 5MB limit
  fileFilter: (req, file, cb) => {
    const allowedTypes = /jpeg|jpg|png|gif|svg/;
    const extname = allowedTypes.test(path.extname(file.originalname).toLowerCase());
    const mimetype = allowedTypes.test(file.mimetype);
    if (extname && mimetype) {
      return cb(null, true);
    }
    cb(new Error('Only image files (jpeg, jpg, png, gif, svg) are allowed'));
  }
});

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

// POST /api/generate-ppt/:id - Generate customized PPTX from campaign brief
router.post('/generate-ppt/:id', upload.single('logo'), async (req, res) => {
  try {
    // Fetch the campaign brief
    const campaignBrief = await CampaignBrief.findById(req.params.id);

    if (!campaignBrief) {
      return res.status(404).json({
        success: false,
        message: 'Campaign brief not found'
      });
    }

    // Validate primary channels - focus only on TV, RADIO, DIGITAL
    const validChannels = ['TV', 'RADIO', 'DIGITAL'];
    const clientChannels = campaignBrief.primaryChannels || [];
    const filteredChannels = clientChannels.filter(ch => 
      validChannels.includes(ch.toUpperCase())
    );

    if (filteredChannels.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'At least one valid primary channel (TV, RADIO, or DIGITAL) is required'
      });
    }

    // Format data for PPTX
    const pptxData = formatCampaignBriefForPptx(campaignBrief);

    // Get logo path if uploaded
    const logoPath = req.file ? req.file.path : null;

    // Generate customized PPTX
    const outputPath = await generateCustomPptx(pptxData, logoPath);

    // Send the file for download
    const filename = path.basename(outputPath);
    res.download(outputPath, filename, (err) => {
      if (err) {
        console.error('Error sending file:', err);
      }
      // Optionally cleanup the generated file after download
      // fs.unlinkSync(outputPath);
    });

  } catch (error) {
    if (error.name === 'CastError') {
      return res.status(400).json({
        success: false,
        message: 'Invalid campaign brief ID'
      });
    }

    console.error('Error generating PPTX:', error);
    res.status(500).json({
      success: false,
      message: 'An error occurred while generating the presentation.'
    });
  }
});

// POST /api/generate-ppt-direct - Generate PPTX with direct input (no DB)
router.post('/generate-ppt-direct', upload.single('logo'), async (req, res) => {
  try {
    const {
      brandName,
      industry,
      targetAgeMin,
      targetAgeMax,
      targetGender,
      primaryChannels,
      keyRegions
    } = req.body;

    // Validate required fields
    if (!brandName || !industry) {
      return res.status(400).json({
        success: false,
        message: 'Brand name and industry are required'
      });
    }

    // Parse arrays if they come as strings
    const parsedChannels = typeof primaryChannels === 'string' 
      ? JSON.parse(primaryChannels) 
      : primaryChannels || [];
    
    const parsedRegions = typeof keyRegions === 'string'
      ? JSON.parse(keyRegions)
      : keyRegions || [];

    // Validate primary channels - focus only on TV, RADIO, DIGITAL
    const validChannels = ['TV', 'RADIO', 'DIGITAL'];
    const filteredChannels = parsedChannels.filter(ch => 
      validChannels.includes(ch.toUpperCase())
    );

    if (filteredChannels.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'At least one valid primary channel (TV, RADIO, or DIGITAL) is required'
      });
    }

    // Format target audience
    const targetAudience = targetGender && targetAgeMin && targetAgeMax
      ? `${targetGender}, Age ${targetAgeMin}-${targetAgeMax}`
      : targetGender || 'All Demographics';

    const pptxData = {
      brandName,
      industry,
      targetAudience,
      primaryChannels: filteredChannels,
      keyRegions: parsedRegions
    };

    // Get logo path if uploaded
    const logoPath = req.file ? req.file.path : null;

    // Generate customized PPTX
    const outputPath = await generateCustomPptx(pptxData, logoPath);

    // Send the file for download
    const filename = path.basename(outputPath);
    res.download(outputPath, filename, (err) => {
      if (err) {
        console.error('Error sending file:', err);
      }
    });

  } catch (error) {
    console.error('Error generating PPTX:', error);
    res.status(500).json({
      success: false,
      message: 'An error occurred while generating the presentation.'
    });
  }
});

module.exports = router;

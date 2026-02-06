const fs = require('fs');
const path = require('path');
const archiver = require('archiver');
const unzipper = require('unzipper');
const { v4: uuidv4 } = require('uuid');

/**
 * PPTX Modifier Utility
 * Modifies slide 4 of the EP-Sales-Pitch-Deck.pptx template
 * with client-specific data
 */

const TEMPLATE_PATH = path.join(__dirname, '..', 'input-client-ppt', 'EP-Sales-Pitch-Deck.pptx');
const OUTPUT_DIR = path.join(__dirname, '..', 'output-ppt');
const TEMP_DIR = path.join(__dirname, '..', 'temp-extract');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

/**
 * Extract PPTX to temporary folder
 */
async function extractPptx(pptxPath, extractPath) {
  return new Promise((resolve, reject) => {
    fs.createReadStream(pptxPath)
      .pipe(unzipper.Extract({ path: extractPath }))
      .on('close', resolve)
      .on('error', reject);
  });
}

/**
 * Repack extracted folder into PPTX
 */
async function repackPptx(extractPath, outputPath) {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outputPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', resolve);
    archive.on('error', reject);

    archive.pipe(output);
    archive.directory(extractPath, false);
    archive.finalize();
  });
}

/**
 * Modify slide 4 XML with client data
 * 
 * Fields to update:
 * - Brand Name (TextBox 13)
 * - Industry (TextBox 13)
 * - Target Audience (TextBox 12) - age range + gender
 * - Primary Channel (TextBox 12) - TV, RADIO, DIGITAL
 * - Key Regions (TextBox 12)
 */
function modifySlide4Xml(xmlContent, clientData) {
  let modifiedXml = xmlContent;

  const {
    brandName = 'Brand Name',
    industry = 'Industry',
    targetAudience = 'All Demographics',
    primaryChannels = ['TV'],
    keyRegions = []
  } = clientData;

  // Format primary channels for display (focus on TV, RADIO, DIGITAL)
  const validChannels = primaryChannels.filter(ch => 
    ['TV', 'RADIO', 'DIGITAL'].includes(ch.toUpperCase())
  );
  const primaryChannelText = validChannels.length > 0 
    ? validChannels.join(', ') 
    : primaryChannels.join(', ');

  // Format key regions for display
  const keyRegionsText = Array.isArray(keyRegions) && keyRegions.length > 0
    ? keyRegions.join(', ')
    : 'Pan India';

  // ====== Update TextBox 13 (Brand Name & Industry) ======
  // Original: <a:t>Brand Name: </a:t>...<a:t>Go Colors</a:t>
  // Original: <a:t>Industry:</a:t>...<a:t> Clothing</a:t>
  
  // Replace "Go Colors" with new brand name
  modifiedXml = modifiedXml.replace(
    /(<a:rPr lang="en-US" sz="3799"><a:solidFill><a:srgbClr val="510C3C"\/><\/a:solidFill><a:latin typeface="Geometria"\/><a:ea typeface="Geometria"\/><a:cs typeface="Geometria"\/><a:sym typeface="Geometria"\/><\/a:rPr><a:t>)Go Colors(<\/a:t>)/,
    `$1${escapeXml(brandName)}$2`
  );

  // Replace "Clothing" with new industry
  modifiedXml = modifiedXml.replace(
    /(<a:rPr lang="en-US" sz="3799"><a:solidFill><a:srgbClr val="510C3C"\/><\/a:solidFill><a:latin typeface="Geometria"\/><a:ea typeface="Geometria"\/><a:cs typeface="Geometria"\/><a:sym typeface="Geometria"\/><\/a:rPr><a:t>) Clothing(<\/a:t>)/,
    `$1 ${escapeXml(industry)}$2`
  );

  // ====== Update TextBox 12 (Target Audience, Primary Channel, Key Regions) ======
  
  // Replace "Female" (target audience)
  modifiedXml = modifiedXml.replace(
    /(<a:rPr lang="en-US" sz="3000"><a:solidFill><a:srgbClr val="000000"\/><\/a:solidFill><a:latin typeface="Geometria"\/><a:ea typeface="Geometria"\/><a:cs typeface="Geometria"\/><a:sym typeface="Geometria"\/><\/a:rPr><a:t>) Female (<\/a:t>)/,
    `$1 ${escapeXml(targetAudience)} $2`.replace('$2', '</a:t>')
  );

  // More robust replacement for target audience
  modifiedXml = modifiedXml.replace(
    /> Female <\/a:t>/g,
    `> ${escapeXml(targetAudience)} </a:t>`
  );

  // Replace "Cinema" (primary channel)
  modifiedXml = modifiedXml.replace(
    />Cinema<\/a:t>/g,
    `>${escapeXml(primaryChannelText)}</a:t>`
  );

  // Replace key regions
  modifiedXml = modifiedXml.replace(
    /> Chennai, Delhi NCR, Bengaluru, Hyderabad, Pune<\/a:t>/g,
    `> ${escapeXml(keyRegionsText)}</a:t>`
  );

  return modifiedXml;
}

/**
 * Escape special XML characters
 */
function escapeXml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Copy logo image to media folder and update relationships
 */
async function updateLogo(extractPath, logoPath) {
  if (!logoPath || !fs.existsSync(logoPath)) {
    console.log('No logo provided or logo file not found, skipping logo update');
    return;
  }

  const mediaDir = path.join(extractPath, 'ppt', 'media');
  const logoExt = path.extname(logoPath).toLowerCase();
  const newLogoName = `image8${logoExt}`; // rId5 points to image8.png
  const destPath = path.join(mediaDir, newLogoName);

  // Copy the new logo
  fs.copyFileSync(logoPath, destPath);

  // If extension changed, update the relationships file
  if (logoExt !== '.png') {
    const relsPath = path.join(extractPath, 'ppt', 'slides', '_rels', 'slide4.xml.rels');
    let relsContent = fs.readFileSync(relsPath, 'utf8');
    relsContent = relsContent.replace(
      /Target="\.\.\/media\/image8\.png"/,
      `Target="../media/${newLogoName}"`
    );
    fs.writeFileSync(relsPath, relsContent);
  }
}

/**
 * Main function to generate customized PPTX
 * 
 * @param {Object} clientData - Client data from CampaignBrief
 * @param {string} clientData.brandName - Brand name
 * @param {string} clientData.industry - Industry
 * @param {string} clientData.targetAudience - Formatted target audience string
 * @param {string[]} clientData.primaryChannels - Primary channels (TV, RADIO, DIGITAL)
 * @param {string[]} clientData.keyRegions - Key regions
 * @param {string} [logoPath] - Optional path to logo image
 * @returns {Promise<string>} Path to generated PPTX file
 */
async function generateCustomPptx(clientData, logoPath = null) {
  const uniqueId = uuidv4();
  const extractPath = path.join(TEMP_DIR, uniqueId);
  const outputFileName = `${clientData.brandName.replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}.pptx`;
  const outputPath = path.join(OUTPUT_DIR, outputFileName);

  try {
    // Create temp directory
    if (!fs.existsSync(extractPath)) {
      fs.mkdirSync(extractPath, { recursive: true });
    }

    // Extract PPTX
    console.log('Extracting template PPTX...');
    await extractPptx(TEMPLATE_PATH, extractPath);

    // Wait for extraction to complete fully
    await new Promise(resolve => setTimeout(resolve, 500));

    // Read slide 4 XML
    const slide4Path = path.join(extractPath, 'ppt', 'slides', 'slide4.xml');
    let slide4Xml = fs.readFileSync(slide4Path, 'utf8');

    // Modify slide 4
    console.log('Modifying slide 4 with client data...');
    slide4Xml = modifySlide4Xml(slide4Xml, clientData);

    // Write modified slide 4
    fs.writeFileSync(slide4Path, slide4Xml);

    // Update logo if provided
    if (logoPath) {
      console.log('Updating logo...');
      await updateLogo(extractPath, logoPath);
    }

    // Repack PPTX
    console.log('Creating customized PPTX...');
    await repackPptx(extractPath, outputPath);

    // Cleanup temp folder
    fs.rmSync(extractPath, { recursive: true, force: true });

    console.log(`Generated: ${outputPath}`);
    return outputPath;

  } catch (error) {
    // Cleanup on error
    if (fs.existsSync(extractPath)) {
      fs.rmSync(extractPath, { recursive: true, force: true });
    }
    throw error;
  }
}

/**
 * Format CampaignBrief data for PPTX modification
 */
function formatCampaignBriefForPptx(campaignBrief) {
  // Format target audience from age range and gender
  const targetAudience = `${campaignBrief.targetGender}, Age ${campaignBrief.targetAgeMin}-${campaignBrief.targetAgeMax}`;

  return {
    brandName: campaignBrief.brandName,
    industry: campaignBrief.industry,
    targetAudience,
    primaryChannels: campaignBrief.primaryChannels || [],
    keyRegions: campaignBrief.keyRegions || []
  };
}

module.exports = {
  generateCustomPptx,
  formatCampaignBriefForPptx,
  modifySlide4Xml,
  escapeXml
};

const sharp = require('sharp');
const fetch = require('node-fetch');

/**
 * Custom formatter to sum numerical values in an array
 *
 * @exampleContext {"data": {"salaries": [1000, 2000, 3000]}}
 * @example [{d.salaries:customSum}] => 6000
 *
 * @param  {Array} value Array of numbers to sum
 * @return {Number}      Sum of all numerical values
 */
function cSum(value) {
  if (!Array.isArray(value)) {
    throw new Error('customSum formatter expects an array');
  }
  return value.reduce((sum, item) => {
    const num = parseFloat(item);
    return isNaN(num) ? sum : sum + num;
  }, 0);
}

/**
 * Custom formatter for dynamic image insertion
 *
 * Supports URLs or Base64 strings, resizes images to fit specified dimensions
 *
 * @exampleContext {"data": {"image": "https://example.com/sample.png"}}
 * @example [{d.image:imageFit(200,200)}] => Base64-encoded resized image
 *
 * @param  {String} value  Image URL or Base64 string
 * @param  {Number} width  [optional] Target width in pixels
 * @param  {Number} height [optional] Target height in pixels
 * @return {String}        Base64-encoded image (data:image/png;base64,...)
 */
async function imgFit(value, width = 200, height = 200) {
  let buffer;
  try {
    if (value.startsWith('data:image')) {
      // Handle Base64
      const base64Data = value.replace(/^data:image\/\w+;base64,/, '');
      buffer = Buffer.from(base64Data, 'base64');
    } else {
      // Handle URL
      const response = await fetch(value);
      if (!response.ok) throw new Error(`Failed to fetch image: ${response.statusText}`);
      buffer = await response.buffer();
    }

    // Resize image using sharp
    const resized = await sharp(buffer)
      .resize({ width: parseInt(width), height: parseInt(height), fit: 'fill' })
      .png() // Ensure PNG output for compatibility
      .toBuffer();

    return `data:image/png;base64,${resized.toString('base64')}`;
  } catch (error) {
    throw new Error(`imgFit formatter failed: ${error.message}`);
  }
}

module.exports = {
  cSum,
  imgFit
};
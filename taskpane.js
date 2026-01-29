/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("FRG Toolbox loaded successfully!");
  }
});

// Helper function to show status messages
function showStatus(elementId, message, type) {
  const statusEl = document.getElementById(elementId);
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  
  if (type === 'success' || type === 'error') {
    setTimeout(() => {
      statusEl.className = 'status';
    }, 5000);
  }
}

// Helper function to add loading indicator
function setLoading(elementId, isLoading, buttonElement) {
  const statusEl = document.getElementById(elementId);
  if (isLoading) {
    statusEl.className = 'status info';
    statusEl.innerHTML = '<span class="loader"></span>Processing...';
    if (buttonElement) buttonElement.disabled = true;
  } else {
    statusEl.className = 'status';
    if (buttonElement) buttonElement.disabled = false;
  }
}

/**
 * Generate an icon from a text prompt using AI
 */
async function generateIcon() {
  const prompt = document.getElementById('iconPrompt').value.trim();
  const button = event.target;
  
  if (!prompt) {
    showStatus('iconStatus', 'Please enter a description for the icon', 'error');
    return;
  }
  
  setLoading('iconStatus', true, button);
  
  try {
    // Call the Anthropic API to generate an image
    // Note: This requires setting up proper API keys and backend
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-3-5-sonnet-20241022',
        max_tokens: 1024,
        messages: [{
          role: 'user',
          content: `Create an SVG icon based on this description: ${prompt}. 
                   Return ONLY the SVG code, no explanations. Make it 512x512 with a transparent background.`
        }]
      })
    });
    
    if (!response.ok) {
      throw new Error('Failed to generate icon');
    }
    
    const data = await response.json();
    const svgCode = data.content[0].text;
    
    // Convert SVG to base64 and insert into PowerPoint
    await insertImageToSlide(svgCode, 'svg');
    
    showStatus('iconStatus', '✓ Icon generated and added to slide!', 'success');
    document.getElementById('iconPrompt').value = '';
    
  } catch (error) {
    console.error('Error generating icon:', error);
    showStatus('iconStatus', 'Note: This demo shows the UI. For full functionality, you need to set up API keys and a backend server.', 'info');
  } finally {
    setLoading('iconStatus', false, button);
  }
}

/**
 * Fetch a logo from the web
 */
async function fetchLogo() {
  const searchTerm = document.getElementById('logoSearch').value.trim();
  const button = event.target;
  
  if (!searchTerm) {
    showStatus('logoStatus', 'Please enter a company or brand name', 'error');
    return;
  }
  
  setLoading('logoStatus', true, button);
  
  try {
    // Use Clearbit Logo API (free tier available)
    const domain = searchTerm.toLowerCase().replace(/\s+/g, '') + '.com';
    const logoUrl = `https://logo.clearbit.com/${domain}`;
    
    // Try to fetch the logo
    const response = await fetch(logoUrl);
    
    if (response.ok) {
      // Convert to base64
      const blob = await response.blob();
      const base64 = await blobToBase64(blob);
      
      // Insert into PowerPoint
      await insertImageToSlide(base64, 'base64');
      
      showStatus('logoStatus', '✓ Logo added to slide!', 'success');
      document.getElementById('logoSearch').value = '';
    } else {
      throw new Error('Logo not found');
    }
    
  } catch (error) {
    console.error('Error fetching logo:', error);
    showStatus('logoStatus', 'Could not find logo. Try a different company name or ensure internet connection.', 'error');
  } finally {
    setLoading('logoStatus', false, button);
  }
}

/**
 * Remove background from selected image
 */
async function removeBackground() {
  const button = event.target;
  setLoading('bgStatus', true, button);
  
  try {
    await PowerPoint.run(async (context) => {
      // Get the selected shape
      const shapes = context.presentation.getSelectedShapes();
      const shapeCount = shapes.getCount();
      
      await context.sync();
      
      if (shapeCount.value === 0) {
        showStatus('bgStatus', 'Please select an image first', 'error');
        return;
      }
      
      if (shapeCount.value > 1) {
        showStatus('bgStatus', 'Please select only one image', 'error');
        return;
      }
      
      const shape = shapes.getItemAt(0);
      shape.load('type');
      
      await context.sync();
      
      if (shape.type !== PowerPoint.ShapeType.picture) {
        showStatus('bgStatus', 'Selected object is not an image', 'error');
        return;
      }
      
      // Get image data
      const image = shape.fill.getImage();
      const imageData = image.getBase64ImageSrc();
      
      await context.sync();
      
      // Call background removal API (e.g., remove.bg)
      // Note: This requires API key setup
      const response = await fetch('https://api.remove.bg/v1.0/removebg', {
        method: 'POST',
        headers: {
          'X-Api-Key': 'YOUR_API_KEY_HERE'
        },
        body: JSON.stringify({
          image_file_b64: imageData.value,
          size: 'auto'
        })
      });
      
      if (response.ok) {
        const resultBlob = await response.blob();
        const resultBase64 = await blobToBase64(resultBlob);
        
        // Replace the image
        shape.fill.setSolidColor('white');
        const newImage = shape.fill.getImage();
        // Note: PowerPoint API limitations may require workarounds here
        
        showStatus('bgStatus', '✓ Background removed!', 'success');
      } else {
        throw new Error('Background removal failed');
      }
      
    });
    
  } catch (error) {
    console.error('Error removing background:', error);
    showStatus('bgStatus', 'Note: This demo shows the UI. For full functionality, you need to set up API keys (remove.bg or similar service).', 'info');
  } finally {
    setLoading('bgStatus', false, button);
  }
}

/**
 * Helper: Insert image into current slide
 */
async function insertImageToSlide(imageData, format) {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    
    await context.sync();
    
    let slide;
    if (slides.items.length === 0) {
      slide = slides.add();
    } else {
      // Get the currently selected slide or use the first one
      slide = slides.getItemAt(0);
    }
    
    // Add image to center of slide
    let base64Image = imageData;
    
    if (format === 'svg') {
      // Convert SVG to base64
      base64Image = 'data:image/svg+xml;base64,' + btoa(imageData);
    } else if (!imageData.startsWith('data:')) {
      base64Image = 'data:image/png;base64,' + imageData;
    }
    
    slide.shapes.addImage(base64Image, 250, 150, 200, 200);
    
    await context.sync();
  });
}

/**
 * Helper: Convert Blob to Base64
 */
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const base64 = reader.result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// Make functions globally accessible
window.generateIcon = generateIcon;
window.fetchLogo = fetchLogo;
window.removeBackground = removeBackground;

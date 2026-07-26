const SAMPLE_TEXT = `Slide 1: Welcome
- Thank you for joining us
- This is a sample presentation
- Created with PPT Creator

Slide 2: Our Mission
- Deliver innovative solutions
- Empower businesses worldwide
- Drive meaningful results

Slide 3: Key Features
- Easy to use interface
- Multiple format options
- Instant download

Slide 4: Getting Started
- Enter your slide content
- Choose a format style
- Click Generate then Download

Slide 5: Thank You
- Questions are welcome
- Contact us anytime`;

const GENERATE_BTN_HTML = `
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
    <path d="M12 5v14M5 12h14"/>
  </svg>
  Generate Presentation`;

let slides = [];
let currentSlideIndex = 0;
let selectedFormat = 'image';
let generatedBlob = null;
let customFormatData = null;
let backgroundImageData = null;
let sampleImageDefault = null;

// Load default sample image on startup
(function loadSampleImage() {
  const img = new Image();
  img.onload = function() {
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    canvas.getContext('2d').drawImage(img, 0, 0);
    sampleImageDefault = canvas.toDataURL('image/png');
  };
  img.src = 'sample-iare.jpg';
})();

const slideInput = document.getElementById('slideInput');
const formatGrid = document.getElementById('formatGrid');
const previewContainer = document.getElementById('previewContainer');
const previewNav = document.getElementById('previewNav');
const slideCounter = document.getElementById('slideCounter');
const prevBtn = document.getElementById('prevSlide');
const nextBtn = document.getElementById('nextSlide');
const generateBtn = document.getElementById('generateBtn');
const downloadBtn = document.getElementById('downloadBtn');
const loadSampleBtn = document.getElementById('loadSample');
const clearTextBtn = document.getElementById('clearText');
const uploadCard = document.getElementById('uploadCard');
const imageCard = document.getElementById('imageCard');
const pptxFileInput = document.getElementById('pptxFileInput');
const imageFileInput = document.getElementById('imageFileInput');
const uploadStatus = document.getElementById('uploadStatus');

function resetDownload() { generatedBlob = null; downloadBtn.disabled = true; }
function markDirty() { resetDownload(); }

function parseSlides(text) {
  const result = [];
  const lines = text.split('\n');
  let current = null;
  for (const raw of lines) {
    const line = raw.trim();
    const slideMatch = line.match(/^Slide\s+(\d+)\s*:\s*(.+)/i);
    if (slideMatch) {
      if (current) result.push(current);
      current = { number: parseInt(slideMatch[1]), title: slideMatch[2].trim(), bullets: [] };
    } else if (current && line.length > 0) {
      // Capture ALL non-empty lines as bullet content
      // Strip leading dash/bullet marker if present, keep the rest as-is
      const clean = line.replace(/^[-*•–—]\s*/, '');
      current.bullets.push(clean);
    }
  }
  if (current) result.push(current);
  return result;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function hexToRgb(hex) {
  hex = hex.replace('#', '');
  if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
  return { r: parseInt(hex.substring(0,2),16), g: parseInt(hex.substring(2,4),16), b: parseInt(hex.substring(4,6),16) };
}

function isLightColor(hex) {
  try {
    const { r, g, b } = hexToRgb(hex);
    return (r * 299 + g * 587 + b * 114) / 1000 > 128;
  } catch { return true; }
}

function isValidHex(h) {
  return h && /^[0-9A-F]{6}$/i.test(h);
}

function makeHex(r, g, b) {
  return [r,g,b].map(x => parseInt(x).toString(16).padStart(2,'0')).join('').toUpperCase();
}

// ============ PPTX PARSING ============

function extractFontFromTag(xml, majorMinor) {
  const block = xml.match(new RegExp('<a:' + majorMinor + 'Font>[\\s\\S]*?</a:' + majorMinor + 'Font>', 'i'));
  if (!block) return null;
  const latin = block[0].match(/<a:latin typeface="([^"]+)"/);
  if (latin) return latin[1];
  const ea = block[0].match(/<a:ea typeface="([^"]+)"/);
  if (ea) return ea[1];
  return null;
}

function buildThemeColorMap(themeXml) {
  const map = {};
  const clrSchemeMatch = themeXml.match(/<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/);
  if (!clrSchemeMatch) return map;
  const cs = clrSchemeMatch[1];

  const tags = ['dk1','dk2','lt1','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
  for (const tag of tags) {
    const re = new RegExp('<a:' + tag + '>[\\s\\S]*?</a:' + tag + '>', 'i');
    const block = cs.match(re);
    if (!block) continue;
    const srgb = block[0].match(/<a:srgbClr val="([0-9A-Fa-f]{6})"\/>/);
    if (srgb) { map[tag] = srgb[1].toUpperCase(); continue; }
    const sysClr = block[0].match(/<a:sysClr lastClr="([0-9A-Fa-f]{6})"/);
    if (sysClr) { map[tag] = sysClr[1].toUpperCase(); continue; }
  }
  return map;
}

function resolveColorInXml(xml, themeMap) {
  if (!xml) return null;
  const srgb = xml.match(/<a:srgbClr val="([0-9A-Fa-f]{6})"\/>/);
  if (srgb) return srgb[1].toUpperCase();
  const schemeRef = xml.match(/<a:schemeClr val="([^"]+)"/);
  if (schemeRef && themeMap[schemeRef[1]]) return themeMap[schemeRef[1]];
  const sysClr = xml.match(/<a:sysClr lastClr="([0-9A-Fa-f]{6})"/);
  if (sysClr) return sysClr[1].toUpperCase();
  return null;
}

function findColorAnywhere(xml, themeMap) {
  const srgb = xml.match(/<a:srgbClr val="([0-9A-Fa-f]{6})"\/>/);
  if (srgb) return srgb[1].toUpperCase();
  const schemeRef = xml.match(/<a:schemeClr val="([^"]+)"/);
  if (schemeRef && themeMap[schemeRef[1]]) return themeMap[schemeRef[1]];
  const sysClr = xml.match(/<a:sysClr lastClr="([0-9A-Fa-f]{6})"/);
  if (sysClr) return sysClr[1].toUpperCase();
  return null;
}

function extractBgColor(bgXml, themeMap) {
  if (!bgXml) return null;
  const solidFill = bgXml.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
  if (solidFill) {
    const c = resolveColorInXml(solidFill[1], themeMap);
    if (c) return c;
  }
  const gradFill = bgXml.match(/<a:gradFill>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"\/>/);
  if (gradFill) return gradFill[1].toUpperCase();
  return findColorAnywhere(bgXml, themeMap);
}

function extractAllTextColors(spXml, themeMap) {
  const colors = [];
  const rPrMatches = [...spXml.matchAll(/<a:rPr[^>]*>([\s\S]*?)<\/a:rPr>/g)];
  for (const [, inner] of rPrMatches) {
    const solidFill = inner.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
    if (solidFill) {
      const c = resolveColorInXml(solidFill[1], themeMap);
      if (c) colors.push(c);
    }
  }
  return colors;
}

function extractAllFontSizes(spXml) {
  const sizes = [];
  const rPrMatches = [...spXml.matchAll(/<a:rPr([^>]*)>/g)];
  for (const [, attrs] of rPrMatches) {
    const szMatch = attrs.match(/sz="(\d+)"/);
    if (szMatch) sizes.push(parseInt(szMatch[1]) / 100);
  }
  return sizes;
}

function extractAllFonts(spXml) {
  const fonts = [];
  const rPrMatches = [...spXml.matchAll(/<a:rPr[^>]*>([\s\S]*?)<\/a:rPr>/g)];
  for (const [, inner] of rPrMatches) {
    const latin = inner.match(/<a:latin typeface="([^"]+)"/);
    if (latin) { fonts.push(latin[1]); continue; }
    const ea = inner.match(/<a:ea typeface="([^"]+)"/);
    if (ea) fonts.push(ea[1]);
  }
  return fonts;
}

function isTitlePlaceholder(spXml) {
  return /<p:ph[^>]*type="(title|ctrTitle)"/i.test(spXml);
}

function isBodyPlaceholder(spXml) {
  return /<p:ph[^>]*type="(body|obj)"/i.test(spXml);
}

async function parsePptxFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);

  const result = {
    bgColor: 'FFFFFF', titleColor: '333333', bodyColor: '555555',
    accentColor: '5B6EEA', accent2Color: null,
    titleFont: 'Arial', bodyFont: 'Arial',
    titleFontSize: 28, bodyFontSize: 16,
  };

  // 1) Theme
  const themeFiles = Object.keys(zip.files).filter(f => f.match(/ppt\/theme\/theme\d+\.xml/));
  let themeMap = {};
  if (themeFiles.length > 0) {
    const themeXml = await zip.files[themeFiles[0]].async('text');
    themeMap = buildThemeColorMap(themeXml);

    if (themeMap.lt1) result.bgColor = themeMap.lt1;
    else if (themeMap.lt2) result.bgColor = themeMap.lt2;
    if (themeMap.dk1) result.titleColor = themeMap.dk1;
    if (themeMap.dk2) result.bodyColor = themeMap.dk2;
    else if (themeMap.dk1) result.bodyColor = themeMap.dk1;
    if (themeMap.accent1) result.accentColor = themeMap.accent1;
    if (themeMap.accent2) result.accent2Color = themeMap.accent2;

    const majorFont = extractFontFromTag(themeXml, 'major');
    const minorFont = extractFontFromTag(themeXml, 'minor');
    if (majorFont) result.titleFont = majorFont;
    if (minorFont) result.bodyFont = minorFont;
  }

  // 2) First slide
  const slideFiles = Object.keys(zip.files).filter(f => f.match(/ppt\/slides\/slide\d+\.xml/)).sort();
  if (slideFiles.length > 0) {
    const slideXml = await zip.files[slideFiles[0]].async('text');

    // Background
    const bgBlock = slideXml.match(/<p:bg>([\s\S]*?)<\/p:bg>/);
    if (bgBlock) {
      const bg = extractBgColor(bgBlock[1], themeMap);
      if (bg && bg !== 'FFFFFF') result.bgColor = bg;
    }

    // Shapes
    const spBlocks = [...slideXml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    let foundTitle = false;
    let foundBody = false;

    for (const [, spContent] of spBlocks) {
      const isTitle = isTitlePlaceholder(spContent);
      const isBody = isBodyPlaceholder(spContent);
      const textColors = extractAllTextColors(spContent, themeMap);
      const fontSizes = extractAllFontSizes(spContent);
      const fonts = extractAllFonts(spContent);
      const hasText = spContent.includes('<a:t>');

      if (isTitle && !foundTitle && hasText) {
        if (textColors.length > 0) result.titleColor = textColors[0];
        if (fontSizes.length > 0) result.titleFontSize = Math.round(fontSizes[0]);
        if (fonts.length > 0) result.titleFont = fonts[0];
        foundTitle = true;
      } else if (isBody && !foundBody && hasText) {
        if (textColors.length > 0) result.bodyColor = textColors[0];
        if (fontSizes.length > 0) result.bodyFontSize = Math.round(fontSizes[0]);
        if (fonts.length > 0) result.bodyFont = fonts[0];
        foundBody = true;
      } else if (hasText && textColors.length > 0) {
        // No placeholder type — use first as title, second as body
        if (!foundTitle) {
          result.titleColor = textColors[0];
          if (fontSizes.length > 0) result.titleFontSize = Math.round(fontSizes[0]);
          if (fonts.length > 0) result.titleFont = fonts[0];
          foundTitle = true;
        } else if (!foundBody) {
          result.bodyColor = textColors[0];
          if (fontSizes.length > 0) result.bodyFontSize = Math.round(fontSizes[0]);
          if (fonts.length > 0) result.bodyFont = fonts[0];
          foundBody = true;
        }
      }
      if (foundTitle && foundBody) break;
    }

    // Fallback: if we found NO text colors at all, grab ALL colors from the slide
    if (foundTitle === false) {
      const allColors = [...slideXml.matchAll(/<a:srgbClr val="([0-9A-Fa-f]{6})"\/>/g)]
        .map(m => m[1].toUpperCase())
        .filter(c => c !== '000000' && c !== 'FFFFFF');
      if (allColors.length > 0) {
        result.titleColor = allColors[0];
        if (allColors.length > 1) result.bodyColor = allColors[1];
        else result.bodyColor = allColors[0];
        if (allColors.length > 2) result.accentColor = allColors[2];
      }
    }
  }

  // 3) Layout fallback for bg
  if (result.bgColor === 'FFFFFF') {
    const layoutFiles = Object.keys(zip.files).filter(f => f.match(/ppt\/slideLayouts\/slideLayout\d+\.xml/));
    if (layoutFiles.length > 0) {
      const layoutXml = await zip.files[layoutFiles[0]].async('text');
      const bgBlock = layoutXml.match(/<p:bg>([\s\S]*?)<\/p:bg>/);
      if (bgBlock) {
        const bg = extractBgColor(bgBlock[1], themeMap);
        if (bg) result.bgColor = bg;
      }
    }
  }

  return result;
}

// ============ FORMAT STYLES ============

function getFormatStyles(format) {
  if (format === 'custom' && customFormatData) {
    return {
      bg: customFormatData.bgColor,
      titleColor: customFormatData.titleColor,
      bodyColor: customFormatData.bodyColor,
      accentColor: customFormatData.accentColor,
      accent2Color: customFormatData.accent2Color || customFormatData.accentColor,
      titleFontFace: customFormatData.titleFont,
      bodyFontFace: customFormatData.bodyFont,
      titleFontSize: customFormatData.titleFontSize,
      bodyFontSize: customFormatData.bodyFontSize,
      titleBold: true,
      bulletStyle: 'bullet',
      decorations: 'none',
    };
  }

  const styles = {
    modern: {
      bg: 'F8F9FF', titleColor: '0070C0', bodyColor: '333333', accentColor: '0070C0', accent2Color: '0070C0',
      titleFontFace: 'Arial', bodyFontFace: 'Arial', titleFontSize: 28, bodyFontSize: 16,
      titleBold: true, bulletStyle: 'number', decorations: 'leftBar',
    },
    minimal: {
      bg: 'FFFFFF', titleColor: '0070C0', bodyColor: '444444', accentColor: '0070C0', accent2Color: '0070C0',
      titleFontFace: 'Helvetica Neue', bodyFontFace: 'Helvetica Neue', titleFontSize: 30, bodyFontSize: 16,
      titleBold: true, bulletStyle: 'none', decorations: 'none',
    },
    bold: {
      bg: '1A1A2E', titleColor: '0070C0', bodyColor: 'CCCCCC', accentColor: '0070C0', accent2Color: '0070C0',
      titleFontFace: 'Arial Black', bodyFontFace: 'Arial', titleFontSize: 32, bodyFontSize: 16,
      titleBold: true, bulletStyle: 'number', decorations: 'rightPanel',
    },
    corporate: {
      bg: 'FFFFFF', titleColor: '0070C0', bodyColor: '333333', accentColor: '0070C0', accent2Color: '0070C0',
      titleFontFace: 'Arial', bodyFontFace: 'Arial', titleFontSize: 28, bodyFontSize: 15,
      titleBold: true, bulletStyle: 'bullet', decorations: 'topBottomBar',
    },
    creative: {
      bg: '0070C0', titleColor: 'FFFFFF', bodyColor: 'FFFFFF', accentColor: 'FFFFFF', accent2Color: 'FFFFFF',
      titleFontFace: 'Arial', bodyFontFace: 'Arial', titleFontSize: 30, bodyFontSize: 16,
      titleBold: true, bulletStyle: 'bullet', decorations: 'circles',
    },
    academic: {
      bg: 'F5F5F0', titleColor: '0070C0', bodyColor: '333333', accentColor: '0070C0', accent2Color: '0070C0',
      titleFontFace: 'Georgia', bodyFontFace: 'Times New Roman', titleFontSize: 26, bodyFontSize: 15,
      titleBold: true, bulletStyle: 'number', decorations: 'topLine',
    },
    image: {
      bg: 'FFFFFF', titleColor: '0070C0', bodyColor: '333333', accentColor: '0070C0', accent2Color: '0070C0',
      titleFontFace: 'Arial', bodyFontFace: 'Arial', titleFontSize: 28, bodyFontSize: 16,
      titleBold: true, bulletStyle: 'none', decorations: 'none',
    },
  };
  return styles[format] || styles.modern;
}

function generateSampleBg() {
  const c = document.createElement('canvas');
  c.width = 1920; c.height = 1080;
  const ctx = c.getContext('2d');
  const g = ctx.createLinearGradient(0, 0, 1920, 1080);
  g.addColorStop(0, '#0070c0');
  g.addColorStop(0.5, '#00a0e8');
  g.addColorStop(1, '#e0f0ff');
  ctx.fillStyle = g;
  ctx.fillRect(0, 0, 1920, 1080);
  ctx.fillStyle = 'rgba(255,255,255,0.08)';
  ctx.beginPath(); ctx.arc(1500, 200, 300, 0, Math.PI * 2); ctx.fill();
  ctx.beginPath(); ctx.arc(400, 800, 200, 0, Math.PI * 2); ctx.fill();
  ctx.fillStyle = 'rgba(255,255,255,0.05)';
  ctx.beginPath(); ctx.arc(900, 500, 400, 0, Math.PI * 2); ctx.fill();
  return c.toDataURL('image/png');
}

let defaultSampleBg = null;

// ============ PREVIEW ============

function renderPreview() {
  if (slides.length === 0) {
    previewContainer.innerHTML = `
      <div class="preview-placeholder">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <rect x="2" y="3" width="20" height="14" rx="2"/>
          <line x1="8" y1="21" x2="16" y2="21"/>
          <line x1="12" y1="17" x2="12" y2="21"/>
        </svg>
        <p>Your slides will appear here</p>
      </div>`;
    previewNav.style.display = 'none';
    return;
  }

  previewNav.style.display = 'flex';
  const slide = slides[currentSlideIndex];
  const s = getFormatStyles(selectedFormat);
  const bgHex = isValidHex(s.bg) ? s.bg : 'FFFFFF';
  const titleHex = isValidHex(s.titleColor) ? s.titleColor : '333333';
  const bodyHex = isValidHex(s.bodyColor) ? s.bodyColor : '555555';
  const light = isLightColor(bgHex);

  let textColor = light ? '#' + bodyHex : '#ddd';
  if (light && isLightColor(bodyHex)) textColor = '#' + bodyHex;
  if (!light && !isLightColor(bodyHex)) textColor = '#CCCCCC';

  // Body text — plain paragraphs with hyphens, no bullet markers
  const bodyHtml = slide.bullets.map((b) => {
    return `<p style="color:${textColor};font-size:0.88rem;margin:6px 0;line-height:1.5;font-family:${s.bodyFontFace},sans-serif;">- ${escapeHtml(b)}</p>`;
  }).join('');

  let titleStyle = `color:#${titleHex};font-size:1.3rem;font-weight:700;margin-bottom:20px;font-family:${s.titleFontFace},sans-serif;`;
  if (s.decorations === 'topBottomBar') titleStyle += `border-bottom:3px solid #${s.accentColor};padding-bottom:10px;`;
  if (s.decorations === 'leftBar' || s.decorations === 'topLine') titleStyle += `border-left:4px solid #${s.accentColor};padding-left:12px;`;

  let bgStyle = `background:#${bgHex};`;
  if (backgroundImageData) {
    bgStyle = `background:#${bgHex} url(${backgroundImageData}) center/cover no-repeat;`;
  } else if (selectedFormat === 'image') {
    const bgData = sampleImageDefault || defaultSampleBg || generateSampleBg();
    bgStyle = `background:#${bgHex} url(${bgData}) center/cover no-repeat;`;
  }

  let fontInfo = '';
  if (selectedFormat === 'custom' && customFormatData) {
    fontInfo = `<div style="position:absolute;bottom:10px;right:16px;font-size:0.7rem;color:${light ? '#888' : '#999'};opacity:0.8;">${s.titleFontFace} / ${s.bodyFontFace} | ${s.titleFontSize}pt / ${s.bodyFontSize}pt</div>`;
  }
  if (selectedFormat === 'image') {
    fontInfo = `<div style="position:absolute;bottom:10px;right:16px;font-size:0.7rem;color:${light ? '#888' : '#999'};opacity:0.8;">Image background</div>`;
  }

  previewContainer.innerHTML = `
    <div class="preview-slide" style="${bgStyle}color:${light ? '#333' : '#fff'};position:relative;">
      <div class="slide-title" style="${titleStyle}">${escapeHtml(slide.title)}</div>
      <div class="slide-body">${bodyHtml}</div>
      ${fontInfo}
    </div>`;

  slideCounter.textContent = `${currentSlideIndex + 1} / ${slides.length}`;
  prevBtn.disabled = currentSlideIndex === 0;
  nextBtn.disabled = currentSlideIndex === slides.length - 1;
}

// ============ GENERATE PPTX ============

async function generatePptx() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'PPT Creator';
  pptx.title = 'Presentation';

  const s = getFormatStyles(selectedFormat);

  slides.forEach((slide, idx) => {
    const presSlide = pptx.addSlide();

    if (backgroundImageData) {
      presSlide.background = { data: backgroundImageData };
    } else if (selectedFormat === 'image') {
      const bgData = sampleImageDefault || defaultSampleBg || generateSampleBg();
      presSlide.background = { data: bgData };
    } else {
      presSlide.background = { fill: s.bg };
    }

    if (selectedFormat !== 'image') {
      if (s.decorations === 'leftBar') {
        presSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.4, h: '100%', fill: { color: s.accentColor } });
      } else if (s.decorations === 'topBottomBar') {
        presSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: s.accentColor } });
        presSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.42, w: '100%', h: 0.08, fill: { color: s.accentColor } });
      } else if (s.decorations === 'rightPanel') {
        presSlide.addShape(pptx.shapes.RECTANGLE, { x: 8, y: 0, w: 5.33, h: '100%', fill: { color: 'FF6B6B', transparency: 85 } });
      } else if (s.decorations === 'circles') {
        presSlide.addShape(pptx.shapes.OVAL, { x: 10.5, y: 0.3, w: 1.2, h: 1.2, fill: { color: 'FFFFFF', transparency: 80 } });
        presSlide.addShape(pptx.shapes.OVAL, { x: 11, y: 3.5, w: 0.8, h: 0.8, fill: { color: 'FFFFFF', transparency: 85 } });
      } else if (s.decorations === 'topLine') {
        presSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 0.05, fill: { color: '2D3436' } });
      }
    }

    const hasLeftAccent = s.decorations === 'leftBar';
    const titleX = hasLeftAccent ? 0.7 : 0.6;
    const titleW = hasLeftAccent ? 11 : 11.5;
    const titleY = s.decorations === 'topBottomBar' ? 0.35 : 0.4;

    presSlide.addText(slide.title, {
      x: titleX, y: titleY, w: titleW, h: 1,
      fontSize: s.titleFontSize, fontFace: s.titleFontFace,
      color: s.titleColor, bold: s.titleBold,
    });

    if (selectedFormat !== 'image') {
      if (s.decorations === 'topBottomBar') {
        presSlide.addShape(pptx.shapes.RECTANGLE, { x: titleX, y: 1.25, w: 2.5, h: 0.04, fill: { color: s.accentColor } });
      } else if (s.decorations === 'topLine') {
        presSlide.addShape(pptx.shapes.RECTANGLE, { x: titleX - 0.15, y: 0.45, w: 0.08, h: 0.7, fill: { color: s.accentColor } });
      }
    }

    const bulletY = 1.55;
    const bodyText = slide.bullets.map(b => `- ${b}`).join('\n');
    presSlide.addText(bodyText, {
      x: titleX, y: bulletY, w: titleW, h: 3.5,
      fontSize: s.bodyFontSize, fontFace: s.bodyFontFace, color: s.bodyColor,
      valign: 'top', lineSpacingMultiple: 1.5,
    });

    presSlide.addText(`${idx + 1}`, {
      x: 11.8, y: 5, w: 0.8, h: 0.4,
      fontSize: 10, color: s.bodyColor, align: 'right', transparency: 50,
    });
  });

  return await pptx.write({ outputType: 'blob' });
}

// ============ DRAG & DROP ============

function setupDragDrop() {
  const panel = document.querySelector('.format-panel');
  ['dragenter', 'dragover'].forEach(evt => {
    panel.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); panel.classList.add('drag-over'); });
  });
  ['dragleave', 'drop'].forEach(evt => {
    panel.addEventListener(evt, (e) => { e.preventDefault(); e.stopPropagation(); panel.classList.remove('drag-over'); });
  });
  panel.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files[0];
    if (!file) return;
    if (file.name.endsWith('.pptx') || file.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
      handleFileUpload(file);
    } else if (file.type.match(/^image\/(png|jpeg|webp)$/)) {
      handleImageUpload(file);
    } else {
      uploadStatus.className = 'upload-status error';
      uploadStatus.innerHTML = 'Please drop a .pptx or image file.';
    }
  });
}

// ============ FILE UPLOAD ============

async function handleFileUpload(file) {
  uploadStatus.className = 'upload-status loading';
  uploadStatus.innerHTML = 'Parsing ' + escapeHtml(file.name) + '...';

  try {
    customFormatData = await parsePptxFile(file);

    formatGrid.querySelectorAll('.format-card').forEach(c => c.classList.remove('selected'));
    uploadCard.classList.add('selected');
    selectedFormat = 'custom';
    backgroundImageData = null;

    const previewColors = `
      <span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#${customFormatData.bgColor};border:1px solid #555;vertical-align:middle;" title="Background: #${customFormatData.bgColor}"></span>
      <span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#${customFormatData.titleColor};border:1px solid #555;vertical-align:middle;" title="Title: #${customFormatData.titleColor}"></span>
      <span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#${customFormatData.bodyColor};border:1px solid #555;vertical-align:middle;" title="Body: #${customFormatData.bodyColor}"></span>
      <span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#${customFormatData.accentColor};border:1px solid #555;vertical-align:middle;" title="Accent: #${customFormatData.accentColor}"></span>`;

    uploadStatus.className = 'upload-status success';
    uploadStatus.innerHTML = `
      <span class="file-name">${escapeHtml(file.name)}</span>
      ${previewColors}
      <span style="color:var(--text-muted);font-size:0.75rem;">${customFormatData.titleFont} / ${customFormatData.bodyFont} | ${customFormatData.titleFontSize}pt / ${customFormatData.bodyFontSize}pt</span>
      <button class="remove-file" id="removeFileBtn">remove</button>`;

    document.getElementById('removeFileBtn').addEventListener('click', removeUploadedFile);
    resetDownload();
    renderPreview();
  } catch (err) {
    console.error('Parse error:', err);
    uploadStatus.className = 'upload-status error';
    uploadStatus.innerHTML = 'Failed to parse file. Make sure it is a valid .pptx file.';
    pptxFileInput.value = '';
  }
}

function handleImageUpload(file) {
  if (!file.type.match(/^image\/(png|jpeg|webp)$/)) {
    uploadStatus.className = 'upload-status error';
    uploadStatus.innerHTML = 'Please select a PNG, JPG, or WebP image.';
    return;
  }

  const reader = new FileReader();
  reader.onload = function(e) {
    backgroundImageData = e.target.result;

    formatGrid.querySelectorAll('.format-card').forEach(c => c.classList.remove('selected'));
    imageCard.classList.add('selected');
    selectedFormat = 'image';
    customFormatData = null;

    uploadStatus.className = 'upload-status success';
    uploadStatus.innerHTML = `
      <span class="file-name">${escapeHtml(file.name)}</span>
      <span style="color:var(--text-muted);font-size:0.75rem;">${(file.size / 1024).toFixed(0)} KB — used as slide background</span>
      <button class="remove-file" id="removeFileBtn">remove</button>`;

    document.getElementById('removeFileBtn').addEventListener('click', removeUploadedFile);
    resetDownload();
    renderPreview();
  };
  reader.readAsDataURL(file);
}

function removeUploadedFile() {
  customFormatData = null;
  backgroundImageData = null;
  pptxFileInput.value = '';
  imageFileInput.value = '';
  uploadCard.classList.remove('selected');
  imageCard.classList.remove('selected');
  uploadStatus.innerHTML = '';
  const imageCardDefault = formatGrid.querySelector('[data-format="image"]');
  if (imageCardDefault) imageCardDefault.classList.add('selected');
  selectedFormat = 'image';
  resetDownload();
  renderPreview();
}

// ============ EVENT LISTENERS ============

slideInput.addEventListener('input', () => {
  slides = parseSlides(slideInput.value);
  currentSlideIndex = 0;
  markDirty();
  renderPreview();
});

formatGrid.addEventListener('click', (e) => {
  const card = e.target.closest('.format-card');
  if (!card) return;

  if (card.dataset.format === 'custom') {
    pptxFileInput.click();
    return;
  }

  if (card.dataset.format === 'image') {
    imageFileInput.click();
    return;
  }

  formatGrid.querySelectorAll('.format-card').forEach(c => c.classList.remove('selected'));
  card.classList.add('selected');
  selectedFormat = card.dataset.format;
  backgroundImageData = null;
  markDirty();
  renderPreview();
});

pptxFileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  handleFileUpload(file);
});

imageFileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  handleImageUpload(file);
});

prevBtn.addEventListener('click', () => {
  if (currentSlideIndex > 0) { currentSlideIndex--; renderPreview(); }
});

nextBtn.addEventListener('click', () => {
  if (currentSlideIndex < slides.length - 1) { currentSlideIndex++; renderPreview(); }
});

loadSampleBtn.addEventListener('click', () => {
  slideInput.value = SAMPLE_TEXT;
  slides = parseSlides(SAMPLE_TEXT);
  currentSlideIndex = 0;
  markDirty();
  renderPreview();
});

clearTextBtn.addEventListener('click', () => {
  slideInput.value = '';
  slides = [];
  currentSlideIndex = 0;
  resetDownload();
  renderPreview();
});

generateBtn.addEventListener('click', async () => {
  if (slides.length === 0) {
    alert('Please enter slide content first.');
    return;
  }
  generateBtn.disabled = true;
  generateBtn.innerHTML = `
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="animation:spin 1s linear infinite;">
      <path d="M21 12a9 9 0 11-6.219-8.56"/>
    </svg>
    Generating...`;
  try {
    generatedBlob = await generatePptx();
    downloadBtn.disabled = false;
  } catch (err) {
    console.error(err);
    alert('Error generating presentation: ' + err.message);
  } finally {
    generateBtn.disabled = false;
    generateBtn.innerHTML = GENERATE_BTN_HTML;
  }
});

downloadBtn.addEventListener('click', () => {
  if (!generatedBlob) return;
  const url = URL.createObjectURL(generatedBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'presentation.pptx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

setupDragDrop();

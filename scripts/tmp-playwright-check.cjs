const { chromium } = require('playwright');
console.log('playwright-loaded', typeof chromium?.launch);

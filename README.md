# Banana Slides

AI-powered PPT generator with professional layouts.

## Installation

\`\`\`bash
npm install
\`\`\`

## Usage

\`\`\`javascript
const BananaSlides = require('./src/core/banana-slides');

const ppt = new BananaSlides({ title: 'My Presentation' });
ppt.addCoverSlide('Title', 'Subtitle', 'Author');
ppt.addTocSlide([{ num: '01', title: 'Chapter 1' }]);
await ppt.save('./output/presentation.pptx');
\`\`\`

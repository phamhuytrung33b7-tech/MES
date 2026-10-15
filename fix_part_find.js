const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

code = code.replace(/const part = parts\.find\(p => p\.id === item\.partId\);\n\s*const originalPart = item\.originalPartId \? parts\.find\(p => p\.id === item\.originalPartId\) : null;/g, 
  `const partStr = item.partId;\n                            const part = parts.find(p => p.id === partStr || p.name === partStr);\n                            const originalPart = item.originalPartId ? parts.find(p => p.id === item.originalPartId || p.name === item.originalPartId) : null;`);

fs.writeFileSync('App.tsx', code);

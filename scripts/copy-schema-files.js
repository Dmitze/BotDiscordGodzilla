const fs = require('fs');
const path = require('path');

// Function to copy file
function copyFile(src, dest) {
  if (!fs.existsSync(path.dirname(dest))) {
    fs.mkdirSync(path.dirname(dest), { recursive: true });
  }
  fs.copyFileSync(src, dest);
  console.log(`Copied ${src} to ${dest}`);
}

// Copy schema files
try {
  // Copy workspace schema
  copyFile(
    path.join(__dirname, '..', 'src', 'workspace', 'sqlite', 'schema.sql'),
    path.join(__dirname, '..', 'dist', 'workspace', 'sqlite', 'schema.sql')
  );
  
  // Copy search schema
  copyFile(
    path.join(__dirname, '..', 'src', 'search', 'sqlite', 'schema.sql'),
    path.join(__dirname, '..', 'dist', 'search', 'sqlite', 'schema.sql')
  );
  
  console.log('Schema files copied successfully!');
} catch (error) {
  console.error('Error copying schema files:', error);
  process.exit(1);
}
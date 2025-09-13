# 📚 Documentation Publishing Guide

This guide explains how to publish the Godzilla Bot documentation on GitHub Pages.

## 🚀 Publishing to GitHub Pages

### 1. Configure GitHub Pages

1. Go to your repository on GitHub
2. Click on **Settings** tab
3. Scroll down to **Pages** section
4. In **Source**, select:
   - **Branch**: `main` (or `master`)
   - **Folder**: `/docs`
5. Click **Save**

### 2. Access Your Documentation

After configuration, your documentation will be available at:
```
https://[username].github.io/BotDiscordGodzilla/
```

For example:
```
https://dmitze.github.io/BotDiscordGodzilla/
```

## 📁 Documentation Structure

The documentation is organized as follows:

```
docs/
├── index.html          # Main entry point with language selection
├── search.html         # Documentation search interface
├── search-index.json   # Search index data
├── uk/                 # Ukrainian documentation
│   ├── README.md       # Main Ukrainian documentation
│   ├── USER_GUIDE.md   # Ukrainian user guide
│   ├── SECURITY.md     # Ukrainian security guide
│   ├── RAG_GUIDE.md    # Ukrainian RAG guide
│   ├── ARCHITECTURE.md # Ukrainian architecture guide
│   ├── COMMANDS_REFERENCE.md # Ukrainian commands reference
│   ├── SETUP.md        # Ukrainian setup guide
│   ├── FAQ_SUPPORT.md  # Ukrainian FAQ and support
│   └── QUICK_START.md  # Ukrainian quick start guide
├── en/                 # English documentation
│   ├── README.md       # Main English documentation
│   ├── USER_GUIDE.md   # English user guide
│   ├── SECURITY.md     # English security guide
│   ├── RAG_GUIDE.md    # English RAG guide
│   ├── ARCHITECTURE.md # English architecture guide
│   ├── COMMANDS_REFERENCE.md # English commands reference
│   ├── SETUP.md        # English setup guide
│   ├── FAQ_SUPPORT.md  # English FAQ and support
│   └── QUICK_START.md  # English quick start guide
└── ...                 # Other documentation files
```

## 🌐 Navigation

Users can navigate the documentation in two ways:

### 1. Main Entry Point
- Visit the root URL to see the language selection page
- Choose between Ukrainian and English documentation

### 2. Direct Access
- Ukrainian documentation: `/uk/README.md`
- English documentation: `/en/README.md`
- Search functionality: `/search.html`

## 🔧 Updating Documentation

To update the published documentation:

1. Make changes to the files in the `docs/` folder
2. Commit and push to your repository
3. GitHub Pages will automatically update (may take a few minutes)

## 🎨 Custom Domain (Optional)

To use a custom domain:

1. In your repository **Settings** → **Pages**
2. In the **Custom domain** field, enter your domain
3. Click **Save**
4. Configure your domain DNS to point to GitHub Pages

## 🔒 HTTPS

GitHub Pages automatically provides HTTPS for all repositories. Your documentation will be served securely.

## 📈 Monitoring

You can monitor your GitHub Pages site:

1. GitHub Pages automatically builds on each push
2. Check build status in **Settings** → **Pages**
3. View build history in **Actions** tab

## 🆘 Troubleshooting

### Common Issues

1. **Page not found (404)**:
   - Ensure GitHub Pages is enabled in Settings
   - Check that files are in the `/docs` folder
   - Verify branch and folder settings

2. **Changes not appearing**:
   - Wait a few minutes for GitHub Pages to rebuild
   - Check the build status in Settings
   - Ensure you've pushed to the correct branch

3. **Broken links**:
   - Verify all relative links point to correct files
   - Check that referenced files exist
   - Test links locally before pushing

### Local Testing

To test documentation locally:

1. Navigate to the `docs/` folder
2. Start a local server:
   ```bash
   # If you have Python installed
   python -m http.server 8000
   
   # Or with Node.js
   npx serve
   
   # Or with PHP
   php -S localhost:8000
   ```
3. Open `http://localhost:8000` in your browser

## 📞 Support

For issues with GitHub Pages:
- [GitHub Pages Documentation](https://docs.github.com/en/pages)
- [GitHub Support](https://support.github.com/)
# 📋 Complete List of Discord Bot Commands with AI

## 🤖 AI Commands (New)

### /ai-search
Description: Natural language search using AI
Usage: `/ai-search query:"find all documents about personnel"`
Examples: 
- `/ai-search query:"find orders from January 2024"`
- `/ai-search query:"show me equipment maintenance records"`

### /ai-analyze
Description: Automatic AI data analysis from table
Usage: `/ai-analyze`
Result: Statistics, trends, anomalies, recommendations

### /ai-recommend
Description: AI recommendations based on data
Usage: `/ai-recommend`
Result: Optimization tips, efficiency improvements

### /ai-report
Description: Creates intelligent report
Usage: `/ai-report type:<report_type>`
Report Types:
- `general-report` - general report
- `stock-report` - stock report
- `sales-report` - sales report
- `suppliers-report` - suppliers report

## 📊 Main Commands (Existing)

### /summary
Description: Shows summary values from table
Usage: `/summary`
Result: Total cost, quantity, price

### /recent
Description: Shows last 10 records from table
Usage: `/recent`
Result: Last records from table

### /search
Description: Search by table fields
Usage: `/search field:<field> query:<query>`
Fields for search: Name, Serial Number, Counterparty, Quantity, Price

### /advanced-search
Description: Search by multiple fields
Usage: `/advanced-search`
Parameters:
- Search by product name
- Search by counterparty
- Search by serial number
- Show products more expensive than this value
- Show products with quantity more than

### /export-search
Description: Exports search results to Excel
Usage: `/export-search field:<field> query:<query>`
Result: Excel file with search results

### /export-all
Description: Exports entire table to Excel
Usage: `/export-all`
Result: Excel file with entire table

### /help
Description: Shows list of all available commands
Usage: `/help`
Result: List of commands with description

## 💬 Text Commands (Existing)

### /add-record
Description: Adds new record via Google Apps Script
Usage: `/add-record`
Result: New record added to table

### /export-table
Description: Exports entire table
Usage: `/export-table`
Result: Excel file with entire table

## 🎯 Usage Examples

### AI Search (Most Useful Examples):

Combined search: `/ai-search query:"find all equipment from supplier ABC with price above 10000"`

Regular search: `/ai-search query:"show me all maintenance records"`

Smart search: `/ai-search query:"which equipment needs maintenance this month"`

### AI Reports (Examples):

General report: `/ai-report type:general-report`

Stock report: `/ai-report type:stock-report`

Sales report: `/ai-report type:sales-report`

Suppliers report: `/ai-report type:suppliers-report`

## 🔧 Technical Details

### Limitations:
- AI-search: maximum 20 results
- AI-analysis: first 50 data rows
- File export: automatic deletion after 10 seconds

### Requirements:
- OpenAI API key for AI functions
- Google Sheets API for working with tables
- Discord Bot Token for bot operation

### Caching:
- Search results are cached for 5 minutes
- Cached results can be exported

## 🆘 Help

If commands don't work:

1. Check bot rights on server
2. Ensure all API keys are configured
3. Check logs in folder

For AI functions:
1. Ensure OpenAI API key exists
2. Run test: 
3. Check OpenAI account balance

For Google Sheets:
1. Check access to table
2. Ensure Google Sheets API is enabled
3. Verify SHEET_ID correctness

💡 Tip: Start with `/help` command to familiarize yourself with all bot capabilities!
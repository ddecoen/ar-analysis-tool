# AR Analysis Tool 📊

Automated Accounts Receivable aging and collections analysis with smart exclusions for wire fees, tax withholdings, and other non-collectible items.

## Features ✨

- **Automated Days Past Due Calculation**: Correctly handles paid vs unpaid invoices
- **Smart Exclusions**: Automatically excludes wire fees and tax withholdings
- **Professional Excel Reports**: Executive summary, detailed data, and aging analysis
- **Configurable**: Easy to customize thresholds and exclusions
- **Audit Ready**: Clean categorization for financial reporting

## Quick Start 🚀

### Option 1: Command Line Usage
```bash
# Install requirements
pip install -r requirements.txt

# Run analysis
python ar_analysis.py your_invoice_data.xlsx output_report.xlsx
```

### Option 2: Import as Module
```python
from ar_analysis import ARAnalyzer

# Create analyzer
analyzer = ARAnalyzer('invoices.xlsx', 'ar_report.xlsx')

# Run analysis
analyzer.run_analysis()
```

## Input File Format 📋

Your Excel file should contain these columns (column names can be mapped in config.ini):
- **Document Number**: Invoice/document identifier
- **Name**: Customer name
- **Invoice Date**: When invoice was created
- **Due Date**: When payment is due
- **Payment Date**: When payment was received (blank if unpaid)
- **Amount**: Invoice amount

## Output Reports 📈

The tool generates a comprehensive Excel report with three sheets:

### 1. Executive Summary
- Key AR metrics and collection rates
- Aging breakdown with percentages
- Key findings and recommendations
- Perfect for management presentations

### 2. Invoice Data
- All invoice details sorted by days past due
- Proper date and currency formatting
- Notes column identifying exclusions
- Clean, audit-ready format

### 3. Collections Analysis
- Detailed aging buckets
- Amount and count breakdowns
- Percentage analysis

## Configuration ⚙️

Customize the analysis by editing `config.ini`:

```ini
[Exclusions]
wire_fee_threshold = 100
india_withholding_docs = 3148

[Column_Mapping]
Your Column Name = Standard Name
```

## Smart Exclusions 🎯

The tool automatically identifies and excludes:

1. **Wire Fees**: Invoices ≤ $100 (configurable)
   - Note: "Remaining wire fees - excluded from AR"

2. **Tax Withholdings**: Specific document numbers
   - Note: "India Withholding Tax - excluded from AR"

3. **Custom Exclusions**: Add your own via configuration

## Key Metrics Calculated 📊

- **Collectible AR Balance**: True receivables excluding non-collectible items
- **Collection Rate**: Percentage of invoices successfully collected
- **Aging Analysis**: Current, 1-30, 31-60, 61-90, 90+ days past due
- **Risk Assessment**: High-risk accounts requiring attention

## Use Cases 💼

- **Monthly AR Reviews**: Regular collection performance analysis
- **Audit Preparation**: Clean, categorized AR for external audits
- **Board Reporting**: Executive-level AR summaries
- **Collection Management**: Identify priority accounts for follow-up
- **Fundraising**: Professional AR analysis for investor due diligence

## Customization Examples 🔧

### Add New Exclusion Type
```python
# In ar_analysis.py, modify the categorize_invoices method
elif doc_num in self.bad_debt_docs:
    self.df.at[idx, 'Category'] = 'Excluded'
    self.df.at[idx, 'Exclusion Reason'] = 'Bad debt write-off'
```

### Change Aging Buckets
```python
# Modify the categorize_aging function
if days_past_due == 0:
    return 'Current'
elif days_past_due <= 15:  # Custom 15-day bucket
    return '1-15 Days Past Due'
# ... etc
```

### Custom Wire Fee Logic
```python
# More sophisticated wire fee detection
if amount <= 100 and 'wire' in str(row['Name']).lower():
    # Mark as wire fee
```

## Integration Options 🔗

### GitHub Repository
1. Create new repository on GitHub
2. Upload these files:
   - `ar_analysis.py`
   - `requirements.txt` 
   - `config.ini`
   - `README.md`

### Automation Ideas
- **Scheduled Reports**: Use cron/Task Scheduler to run monthly
- **Email Integration**: Auto-send reports to management
- **Database Integration**: Pull data from ERP systems
- **API Integration**: Connect to NetSuite, QuickBooks, etc.

## Troubleshooting 🔧

### Common Issues:
1. **Column Name Mismatch**: Update `config.ini` column mapping
2. **Date Format Issues**: Ensure dates are in Excel date format
3. **Missing Data**: Check for required columns in input file

### Error Messages:
- "Column not found": Update column mapping in config
- "No data loaded": Check file path and Excel file format
- "Permission denied": Close Excel file if open

## Advanced Usage 🚀

### Batch Processing Multiple Files
```python
import glob
from ar_analysis import ARAnalyzer

# Process all Excel files in a directory
for file in glob.glob("*.xlsx"):
    analyzer = ARAnalyzer(file)
    analyzer.run_analysis()
```

### Custom Reporting
```python
analyzer = ARAnalyzer('data.xlsx')
analyzer.run_analysis()

# Access calculated metrics
print(f"AR Balance: ${analyzer.metrics['collectible_ar']:,.2f}")
print(f"Collection Rate: {analyzer.metrics['collection_rate']:.1f}%")
```

## Support & Updates 💬

- **Documentation**: This README and inline code comments
- **Configuration**: `config.ini` for easy customization
- **Extensibility**: Clean, modular code for easy modification

## License 📄

Open source - feel free to modify and distribute for your organization's needs.

---

**Created for Series C fundraising and audit preparation** 🎯
*Automated AR analysis that saves hours of manual work*

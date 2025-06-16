# GTM Documentation Automation

🚀 **Generate complete Google Tag Manager documentation in 1 click using Google Apps Script**

## The Problem
- Manual GTM documentation takes hours
- Tracking constantly changes across team members
- Stakeholders come to you with same question everytime "what does this event track?"
- You become the documentation machine

## The Solution
Automatically extract all GA4 event tags, parameters, triggers, and descriptions from GTM into a formatted Google Sheet report. 
# Potential: You can import the sheet into BigQuery and transform to check your data quality, get alert if tracking is broken, build Looker Studio dashboard for self serv data etc...

## ✨ Features
- 📊 Complete GA4 event documentation
- 🔄 Auto-updates parameter descriptions
- 📝 Tracks event names and descriptions
- 🎯 Maps triggers to tags
- 📋 Built-in variable descriptions

## 📸 Screenshots
![Demo Screenshot](screenshots/demo.png)
*Generated documentation showing events, parameters, and triggers*

## 🚀 Quick Start

### Prerequisites
- Google Sheets access
- Google Tag Manager access
- GTM Advanced Service enabled

### Installation
1. **Create new Google Sheet**
2. **Open Apps Script** (Extensions → Apps Script)
3. **Copy the code** from `Code.gs`
4. **Enable GTM API** (Libraries → Add library → Script ID: `1-rr7_ggPwd2PlGkmfeWk3mz_5s6_VngHy8TF4gPJO9LvCfqvihQ0ZPQJ`)
5. **Save and run** `onOpen()` function

### Setup
1. Run the script from Google Sheets menu: **GTM Reports → Configure GTM Settings**
2. Enter your GTM Account ID, Container ID, and Workspace ID
3. Click **GTM Reports → Generate GA4 Event Tag Report**

## 📊 Generated Documentation Includes
- Tag names and triggers
- Event names with descriptions
- Parameter mappings
- Built-in variable descriptions
- Custom variable notes from GTM

## 🛠️ Advanced Features
- **Parameters Sheet**: Tracks all parameter values and descriptions
- **Events Sheet**: Manages event names and descriptions  
- **Auto-updates**: Syncs changes from GTM to existing documentation
- **Built-in Variables**: Pre-loaded descriptions for GTM built-in variables

## 📖 Documentation
- [Complete Setup Guide](setup-guide.md)
- [Troubleshooting](docs/troubleshooting.md)
- [Advanced Configuration](docs/advanced.md)

## 🤝 Contributing
Found a bug or have a feature request? Please open an issue or submit a pull request!

## 📄 License
MIT License - feel free to use and modify!

## 🙏 Support
If this saved you time, give it a ⭐ and share with other analysts!
Buy me a coffee: 
(Swedish users): Swish 0733801686
🌍 INTERNATIONAL: https://ko-fi.com/chinhjohnson

---
*Built by Chinh Johnson - Technical Digital Analyst tired of manual documentation*

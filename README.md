# GTM Documentation Automation

🚀 **Generate complete Google Tag Manager documentation in 1 click by ultilizing Notes function in GTM and using Google Apps Script**

## The Problem
- Manual GTM documentation takes hours
- Tracking constantly changes across team members
- Stakeholders come to you with same question everytime "what does this event track?"
- You become the documentation machine

## The Solution
Automatically extract all GA4 event tags, parameters, triggers, and descriptions into a formatted Google Sheet report ultilizing GTM Notes function.

# Potential next steps:
You can import the sheet into BigQuery to transform and analyze the data, monitor data quality, get alerts if tracking breaks, and even build a self-serve dashboard in Looker Studio

## ✨ Features
- 📊 Complete GA4 event documentation ultilising Notes features in GTM
- 🔄 Scheduled updates events and parameter descriptions
- 🎯 Maps triggers to tags
- 📋 Built-in variable descriptions

## 📸 Youtube tutorial
https://youtu.be/dkCvhEot1tY

### Prerequisites
- Google Sheets access
- Google Tag Manager access
- Basic understand of Google Tag Manager

### Installation
1. **Create new Google Sheet**
2. **Open Apps Script** (Extensions → Apps Script)
3. **Copy the code** from `Code.gs`
4. **Enable GTM API** (Services → Add Tag Manager API)
5. **Save and run** `onOpen()` function

### Setup
1. Run the script from Google Sheets menu: **GTM Reports → Configure GTM Settings**
2. Enter your GTM Account ID, Container ID, and Workspace ID
3. Click **GTM Reports → Generate GA4 Event Tag Report**


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

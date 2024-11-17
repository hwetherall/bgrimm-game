# BGrimm Business Game Scoring System

An automated scoring system for the BGrimm Business Game workshop, using GPT-4 for response evaluation and Google Sheets for data management.

## Features
- Automated scoring of workshop responses using GPT-4
- Real-time dashboard updates
- Team performance tracking
- Round-by-round statistics
- Automatic hypothesis detection

## Tech Stack
- Node.js
- OpenAI GPT-4 API
- Google Sheets API
- Google Forms (for data collection)

## Setup Instructions

### Prerequisites
- Node.js v20.18.0 or higher
- Google Cloud account with Sheets API enabled
- OpenAI API key

### Configuration
1. Clone the repository
```bash
git clone [repository-url]
cd bgrimm-game
```

2. Install dependencies
```bash
npm install
```

3. Create .env file with:
```
OPENAI_API_KEY=your_openai_api_key
SPREADSHEET_ID=your_google_sheet_id
```

4. Add your Google Cloud credentials:
- Create credentials.json from Google Cloud Console
- Place in project root directory

### Running the Application
```bash
node index.js
```

## System Architecture
- Google Form collects team responses
- Node.js application monitors for new submissions
- GPT-4 scores responses based on criteria
- Dashboard updates automatically
- Real-time statistics and rankings

## Contributing
This is a project for the BGrimm Business Game workshop. Feel free to fork and modify for your own use.

## License
[Choose appropriate license]

## Author
[Your Name]

## Acknowledgments
- Built for BGrimm Business Game workshop
- Uses OpenAI's GPT-4 for response evaluation

// Import required dependencies
import { google } from 'googleapis';
import dotenv from 'dotenv';
import OpenAI from 'openai';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

// Configure environment variables
dotenv.config();

// Get the current file's directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Initialize OpenAI
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// Initialize Google Sheets API
const auth = new google.auth.GoogleAuth({
  keyFile: join(__dirname, 'credentials.json'),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

// Configuration
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// Add headers function
// Add headers function
async function addHeaders() {
    try {
      // First add the score column headers
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Form Responses 1!H1:M1',
        valueInputOption: 'RAW',
        requestBody: {
          values: [[
            'Score - Customer Centricity',
            'Score - Time to Market',
            'Score - Culture',
            'Score - Process Changes',
            'Score - Partner Selection',
            'Total Score'
          ]]
        }
      });
      
      // Then add the total formula to each row
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Form Responses 1!M2',
        valueInputOption: 'USER_ENTERED', // This allows formulas
        requestBody: {
          values: [['=SUM(H2:L2)']]
        }
      });
    } catch (error) {
      console.error('Error adding headers:', error);
    }
  }

// Enhanced scoring prompt
const SCORING_PROMPT = `Please score the following response on a scale of 0 to 3 based on these criteria:

0 points: If the answer is fairly general and short, does not go into details and does not point to specific ideas or actions.

1 point: If the answer is very specific, with details and is at least 2 sentences long but not actionable (e.g., it's not clear what could be done next). The response describes a situation or problem without concrete steps to address it.

2 points: If the answer is very specific, with details and is at least 2 sentences long and actionable (i.e., it provides clear, numbered or specific steps that should be done next).

3 points: If the answer is very specific, actionable, AND explicitly contains the word "hypothesis" followed by a testable assumption and method to validate it.

Important Notes:
- The response must be at least 30 words long to be eligible for a score of 1 or more points
- For 3 points, the word "hypothesis" MUST be present
- For 2 points, there must be clear, actionable steps, not just description
- For 1 point, the answer must be detailed but may lack concrete actions

Please analyze carefully and only respond with a single number (0, 1, 2, or 3).

Response to evaluate: `;

// Function to get new responses from Google Sheets
async function getNewResponses() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Form Responses 1!A:M', // Extended to include score columns
    });
    return response.data.values || [];
  } catch (error) {
    console.error('Error fetching responses:', error);
    return [];
  }
}

// Function to score a response using OpenAI
async function scoreResponse(response) {
  try {
    const completion = await openai.chat.completions.create({
      model: "gpt-4",
      messages: [
        {
          role: "system",
          content: "You are an AI scoring responses for a business game. You should only return a single number (0, 1, 2, or 3) based on the scoring criteria."
        },
        {
          role: "user",
          content: SCORING_PROMPT + response
        }
      ],
      temperature: 0,
    });

    const score = parseInt(completion.choices[0].message.content.trim());
    return isNaN(score) ? 0 : score;
  } catch (error) {
    console.error('Error scoring response:', error);
    return 0;
  }
}

// Function to update score in Google Sheets
async function updateScore(row, column, score) {
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Form Responses 1!${column}${row}`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[score]]
      }
    });
    console.log(`Updated score ${score} in cell ${column}${row}`);
  } catch (error) {
    console.error('Error updating score:', error);
  }
}

// Main function to process responses
async function processResponses() {
  console.log('Checking for new responses...');
  const responses = await getNewResponses();
  
  if (responses.length <= 1) return; // Only header row or empty
  
  // Process each response starting from row 2 (after header)
  for (let i = 1; i < responses.length; i++) {
    const row = responses[i];
    // Start from column C (index 2) which contains the first solution
    for (let j = 2; j <= 6; j++) {
      // Check if there's a response but no score
      if (row[j] && (!row[j + 5] || row[j + 5] === '')) {
        console.log(`Processing response for row ${i + 1}, scenario ${j - 1}`);
        const score = await scoreResponse(row[j]);
        // Calculate the score column (H for first scenario, I for second, etc.)
        const scoreColumn = String.fromCharCode(72 + (j - 2)); // H, I, J, K, L
        await updateScore(i + 1, scoreColumn, score);
      }
    }
  }
}

async function setupDashboard() {
    try {
      let dashboardSheetId;
      
      try {
        const response = await sheets.spreadsheets.get({
          spreadsheetId: SPREADSHEET_ID
        });
        
        const dashboardSheet = response.data.sheets.find(sheet => sheet.properties.title === 'Dashboard');
        
        if (dashboardSheet) {
          dashboardSheetId = dashboardSheet.properties.sheetId;
          console.log('Found existing Dashboard sheet');
        } else {
          const addSheetResponse = await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: {
              requests: [{
                addSheet: {
                  properties: {
                    title: 'Dashboard',
                    gridProperties: {
                      rowCount: 50,
                      columnCount: 12
                    },
                    tabColor: {
                      red: 0.2,
                      green: 0.7,
                      blue: 0.9
                    }
                  }
                }
              }]
            }
          });
          dashboardSheetId = addSheetResponse.data.replies[0].addSheet.properties.sheetId;
        }

        // First, get the number of unique teams
        const teamsResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Form Responses 1!B2:B'
        });
        
        // Get unique team names and filter out empty values
        const uniqueTeams = [...new Set(teamsResponse.data.values?.flat().filter(Boolean) || [])];
        const numberOfTeams = uniqueTeams.length;

        // Create the basic headers
        let dashboardData = [
            ['Team Name', 'Round 1', 'Round 2', 'Round 3', 'Round 4', 'Round 5', 'Total Score', 'Rank']
        ];

        // Add a row for each team
        uniqueTeams.forEach((team, index) => {
            dashboardData.push([
                team,
                `=IFERROR(VLOOKUP("${team}",'Form Responses 1'!B:H,7,FALSE),0)`,
                `=IFERROR(VLOOKUP("${team}",'Form Responses 1'!B:I,8,FALSE),0)`,
                `=IFERROR(VLOOKUP("${team}",'Form Responses 1'!B:J,9,FALSE),0)`,
                `=IFERROR(VLOOKUP("${team}",'Form Responses 1'!B:K,10,FALSE),0)`,
                `=IFERROR(VLOOKUP("${team}",'Form Responses 1'!B:L,11,FALSE),0)`,
                `=SUM(B${index + 2}:F${index + 2})`,
                `=RANK(G${index + 2},$G$2:$G${numberOfTeams + 1})`
            ]);
        });

        // Add the average row
        const lastDataRow = numberOfTeams + 1;
        dashboardData.push([
            'Average',
            `=AVERAGE(B2:B${lastDataRow})`,
            `=AVERAGE(C2:C${lastDataRow})`,
            `=AVERAGE(D2:D${lastDataRow})`,
            `=AVERAGE(E2:E${lastDataRow})`,
            `=AVERAGE(F2:F${lastDataRow})`,
            `=AVERAGE(G2:G${lastDataRow})`,
            ''
        ]);

        // Clear existing content
        await sheets.spreadsheets.values.clear({
            spreadsheetId: SPREADSHEET_ID,
            range: `Dashboard!A1:H${lastDataRow + 1}`
        });

        // Update with new data
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `Dashboard!A1:H${lastDataRow + 1}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: {
                values: dashboardData
            }
        });

        // Update formatting to match the new size
        const formatRequests = {
            requests: [
                // Header formatting
                {
                    repeatCell: {
                        range: {
                            sheetId: dashboardSheetId,
                            startRowIndex: 0,
                            endRowIndex: 1,
                            startColumnIndex: 0,
                            endColumnIndex: 8
                        },
                        cell: {
                            userEnteredFormat: {
                                backgroundColor: { red: 0.2, green: 0.2, blue: 0.2 },
                                textFormat: {
                                    foregroundColor: { red: 1, green: 1, blue: 1 },
                                    bold: true
                                },
                                horizontalAlignment: 'CENTER'
                            }
                        },
                        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                    }
                },
                // Center align all number cells (for team rows)
                {
                    repeatCell: {
                        range: {
                            sheetId: dashboardSheetId,
                            startRowIndex: 1,
                            endRowIndex: lastDataRow,
                            startColumnIndex: 1,
                            endColumnIndex: 8
                        },
                        cell: {
                            userEnteredFormat: {
                                horizontalAlignment: 'CENTER'
                            }
                        },
                        fields: 'userEnteredFormat.horizontalAlignment'
                    }
                },
                // Average row formatting (bold and center-aligned)
                {
                    repeatCell: {
                        range: {
                            sheetId: dashboardSheetId,
                            startRowIndex: lastDataRow,
                            endRowIndex: lastDataRow + 1,
                            startColumnIndex: 0,
                            endColumnIndex: 8
                        },
                        cell: {
                            userEnteredFormat: {
                                backgroundColor: { red: 0.9, green: 0.9, blue: 0.9 },
                                textFormat: { bold: true },
                                horizontalAlignment: 'CENTER'
                            }
                        },
                        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                    }
                },
                // Borders
                {
                    updateBorders: {
                        range: {
                            sheetId: dashboardSheetId,
                            startRowIndex: 0,
                            endRowIndex: lastDataRow + 1,
                            startColumnIndex: 0,
                            endColumnIndex: 8
                        },
                        top: { style: 'SOLID' },
                        bottom: { style: 'SOLID' },
                        left: { style: 'SOLID' },
                        right: { style: 'SOLID' }
                    }
                }
            ]
        };

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: formatRequests
        });

      } catch (error) {
        console.error('Error in spreadsheet operations:', error);
      }
    } catch (error) {
      console.error('Error setting up dashboard:', error);
    }
}

// Add this to your existing testConnection function
async function testConnection() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Form Responses 1!A1:B1',
    });
    console.log('Successfully connected to Google Sheets!');
    console.log('Header row:', response.data.values[0]);
    
    // Add headers if they don't exist
    await addHeaders();
    
    // Set up dashboard
    await setupDashboard();
    
    // Start processing responses
    setInterval(processResponses, 30000);
    processResponses();
  } catch (error) {
    console.error('Error connecting to Google Sheets:', error.message);
  }
}

// Start the application
testConnection();
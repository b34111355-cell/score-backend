require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
app.use(cors());
app.use(express.json());

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

async function appendScore({ playerId, points, note }) {
  const sheets = await getSheetsClient();
  const range = `${process.env.SHEET_NAME}!A:D`;
  const values = [[new Date().toISOString(), playerId, points, note || '']];

  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
}

app.post('/api/score', async (req, res) => {
  const { passcode, playerId, points, note } = req.body;

  if (!passcode || passcode !== process.env.TEAM_PASSCODE) {
    return res.status(401).json({ error: '通行碼錯誤' });
  }
  if (!playerId || points === undefined) {
    return res.status(400).json({ error: '缺少 playerId 或 points' });
  }

  try {
    await appendScore({ playerId, points, note });
    res.json({ ok: true });
  } catch (err) {
    console.error('寫入失敗:', err.message);
    res.status(500).json({ error: '寫入試算表失敗' });
  }
});

app.get('/api/health', (req, res) => {
  res.json({ ok: true, time: new Date().toISOString() });
});

// 接收整場比賽的同步 API
app.post('/api/syncMatch', async (req, res) => {
  const { passcode, matchData } = req.body;

  if (!passcode || passcode !== process.env.TEAM_PASSCODE) {
    return res.status(401).json({ error: '通行碼錯誤' });
  }
  if (!matchData) {
    return res.status(400).json({ error: '缺少 matchData' });
  }

  try {
    const sheets = await getSheetsClient();
    const range = `${process.env.SHEET_NAME}!A:D`;
    
    const values = [[
      new Date().toISOString(),
      matchData.id,
      matchData.points ? matchData.points.length : 0,
      JSON.stringify(matchData)
    ]];

    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.SPREADSHEET_ID,
      range,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    });

    res.json({ ok: true, matchId: matchData.id, revision: matchData.revision });
  } catch (err) {
    console.error('寫入失敗:', err.message);
    res.status(500).json({ error: '寫入試算表失敗' });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Score backend listening on http://localhost:${port}`);
});

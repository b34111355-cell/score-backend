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
// 接收整場比賽的同步 API (拆解為比賽結果與逐分明細)
app.post('/api/syncMatch', async (req, res) => {
  const { passcode, matchData } = req.body;

  if (!passcode || passcode !== process.env.TEAM_PASSCODE) {
    return res.status(401).json({ error: '通行碼錯誤' });
  }
  if (!matchData || !matchData.points) {
    return res.status(400).json({ error: '缺少 matchData' });
  }

  try {
    const sheets = await getSheetsClient();
    const now = new Date().toISOString();
    const matchId = matchData.id;
    const config = matchData.config;
    
    // 處理名字 (如果是單打，後面的名字會是空字串)
    const nameA = config.mode === 'doubles' ? `${config.names[0]}/${config.names[1]}` : config.names[0];
    const nameB = config.mode === 'doubles' ? `${config.names[2]}/${config.names[3]}` : config.names[2];

    // 統計這場比賽的總得分
    let scoreA = 0, scoreB = 0;
    matchData.points.forEach(p => { if (p.winnerTeam === 0) scoreA++; else scoreB++; });
    const winnerTeamName = scoreA > scoreB ? nameA : nameB;

    // ================= 1. 寫入 Match_Summary =================
    const matchSummaryValues = [[
      now,
      matchId,
      config.mode === 'doubles' ? '雙打' : '單打',
      nameA,
      nameB,
      winnerTeamName,
      scoreA,
      scoreB
    ]];

    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.SPREADSHEET_ID,
      range: `Match_Summary!A:H`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: matchSummaryValues },
    });

    // ================= 2. 寫入 Point_Details =================
    const pointValues = [];
    
    for (const p of matchData.points) {
      const isSvTeamA = p.serverIdx < 2;
      const serverName = config.names[p.serverIdx];
      
      // 推算接發方
      let receiverName = '';
      if (config.mode === 'doubles') {
         // 簡化推論：如果A發球，接發是B隊的其中一人；反之亦然
         receiverName = isSvTeamA ? `${config.names[2]}/${config.names[3]}` : `${config.names[0]}/${config.names[1]}`;
      } else {
         receiverName = isSvTeamA ? config.names[2] : config.names[0];
      }
      
      const winnerTeamName = p.winnerTeam === 0 ? 'A隊' : 'B隊';
      const endingPlayerName = config.names[p.endingPlayerIdx] || '未知';
      const isWinner = (p.winnerTeam === 0) === isSvTeamA; // 發球方是否得分
      
      // 得分方式轉中文
      const typeMap = { ace: 'Ace', df: '雙誤', serve_winner: '發球直得', winner: '製勝分', return_winner: '接發製勝', ue: '非受迫失誤', fe: '受迫失誤' };
      const endingTypeName = typeMap[p.endingType] || p.endingType;

      pointValues.push([
        now,
        matchId,
        config.mode === 'doubles' ? '雙打' : '單打',
        serverName,
        receiverName,
        p.firstServeIn ? '是' : '否',
        winnerTeamName,
        endingTypeName,
        endingPlayerName,
        isWinner ? '發球方得分' : '接發方得分'
      ]);
    }

    if (pointValues.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: process.env.SPREADSHEET_ID,
        range: `Point_Details!A:J`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: pointValues },
      });
    }

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

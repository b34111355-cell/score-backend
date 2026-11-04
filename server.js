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

// 健康檢查
app.get('/api/health', (req, res) => {
  res.json({ ok: true, time: new Date().toISOString() });
});

// ================= 1. 接收比賽資料 (覆蓋寫入) =================
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
    const spreadsheetId = process.env.SPREADSHEET_ID;
    const now = new Date().toISOString();
    const matchId = matchData.id;
    const config = matchData.config;

    // 取得分頁 ID
    const spreadsheetInfo = await sheets.spreadsheets.get({ spreadsheetId });
    const pdSheet = spreadsheetInfo.data.sheets.find(s => s.properties.title === 'Point_Details');
    const msSheet = spreadsheetInfo.data.sheets.find(s => s.properties.title === 'Match_Summary');

    if (!pdSheet || !msSheet) {
      throw new Error('找不到 Point_Details 或 Match_Summary 分頁');
    }

    // 清除舊資料 (Point_Details)
    const pdRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Point_Details!A:B' });
    const pdRows = pdRes.data.values || [];
    const pdIndicesToDelete = [];
    for (let i = pdRows.length - 1; i > 0; i--) {
      if (pdRows[i][1] === matchId) {
        pdIndicesToDelete.push({ deleteDimension: { range: { sheetId: pdSheet.properties.sheetId, dimension: 'ROWS', startIndex: i, endIndex: i + 1 } } });
      }
    }
    if (pdIndicesToDelete.length > 0) {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: pdIndicesToDelete } });
    }

    // 清除舊資料 (Match_Summary)
    const msRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Match_Summary!A:B' });
    const msRows = msRes.data.values || [];
    const msIndicesToDelete = [];
    for (let i = msRows.length - 1; i > 0; i--) {
      if (msRows[i][1] === matchId) {
        msIndicesToDelete.push({ deleteDimension: { range: { sheetId: msSheet.properties.sheetId, dimension: 'ROWS', startIndex: i, endIndex: i + 1 } } });
      }
    }
    if (msIndicesToDelete.length > 0) {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: msIndicesToDelete } });
    }

    // 寫入 Match_Summary
    const nameA = config.mode === 'doubles' ? `${config.names[0]}/${config.names[1]}` : config.names[0];
    const nameB = config.mode === 'doubles' ? `${config.names[2]}/${config.names[3]}` : config.names[2];
    let scoreA = 0, scoreB = 0;
    matchData.points.forEach(p => { if (p.winnerTeam === 0) scoreA++; else scoreB++; });
    const winnerTeamName = scoreA > scoreB ? nameA : nameB;

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'Match_Summary!A:H',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[now, matchId, config.mode === 'doubles' ? '雙打' : '單打', nameA, nameB, winnerTeamName, scoreA, scoreB]] },
    });

    // 寫入 Point_Details
    const pointValues = [];
    for (const p of matchData.points) {
      const isSvTeamA = p.serverIdx < 2;
      const serverName = config.names[p.serverIdx];
      const receiverName = config.mode === 'doubles' ? (isSvTeamA ? `${config.names[2]}/${config.names[3]}` : `${config.names[0]}/${config.names[1]}`) : (isSvTeamA ? config.names[2] : config.names[0]);
      const winnerTeamName = p.winnerTeam === 0 ? 'A隊' : 'B隊';
      const endingPlayerName = config.names[p.endingPlayerIdx] || '未知';
      const isWinner = (p.winnerTeam === 0) === isSvTeamA;
      const typeMap = { ace: 'Ace', df: '雙誤', serve_winner: '發球直得', winner: '製勝分', return_winner: '接發製勝', ue: '非受迫失誤', fe: '受迫失誤' };

      pointValues.push([now, matchId, config.mode === 'doubles' ? '雙打' : '單打', serverName, receiverName, p.firstServeIn ? '是' : '否', winnerTeamName, typeMap[p.endingType] || p.endingType, endingPlayerName, isWinner ? '發球方得分' : '接發方得分']);
    }

    if (pointValues.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: 'Point_Details!A:J',
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: pointValues },
      });
    }

    res.json({ ok: true, matchId: matchData.id, revision: matchData.revision });
  } catch (err) {
    console.error('🔴 寫入失敗原因:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ================= 2. 讀取所有歷史資料 (給前端算排名) =================
app.get('/api/allData', async (req, res) => {
  const { passcode } = req.query;
  if (!passcode || passcode !== process.env.TEAM_PASSCODE) {
    return res.status(401).json({ error: '通行碼錯誤' });
  }

  try {
    const sheets = await getSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const msRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Match_Summary!A:H' });
    const msRows = msRes.data.values || [];
    const matches = msRows.slice(1).map(row => ({
      time: row[0], matchId: row[1], mode: row[2], teamA: row[3], teamB: row[4], winner: row[5],
      scoreA: parseInt(row[6]) || 0, scoreB: parseInt(row[7]) || 0
    }));

    const pdRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Point_Details!A:J' });
    const pdRows = pdRes.data.values || [];
    const points = pdRows.slice(1).map(row => ({
      time: row[0], matchId: row[1], mode: row[2], server: row[3], receiver: row[4], firstServeIn: row[5] === '是',
      winnerTeam: row[6], endingType: row[7], player: row[8], pointWinner: row[9]
    }));

    res.json({ ok: true, matches, points });
  } catch (err) {
    console.error('讀取失敗:', err.message);
    res.status(500).json({ error: '讀取試算表失敗' });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Score backend listening on http://localhost:${port}`);
});
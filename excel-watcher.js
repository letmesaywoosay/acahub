const fs = require('fs');
const path = require('path');
const chokidar = require('chokidar');
const XLSX = require('xlsx');

// 엑셀 파일 및 빌드 출력 대상 설정
const EXCEL_FILE = path.join(__dirname, '솔루션실습 접속정보 관리.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'data.js');

console.log('==================================================');
console.log('  ACADEMY Hub Excel Watcher 기동 (Trainee View)');
console.log('==================================================');
console.log('감시 파일:', EXCEL_FILE);
console.log('출력 파일:', OUTPUT_FILE);


function formatExcelDate(val) {
  if (!val) return "";
  const num = Number(val);
  
  // 엑셀 날짜 일련번호 범위인 경우 (예: 46210)
  if (!isNaN(num) && num > 40000 && num < 60000) {
    const utc_days  = Math.floor(num - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    // 타임존 보정 (KST 현지 표준시 기준)
    const tzOffset = date_info.getTimezoneOffset() * 60 * 1000;
    const local_date = new Date(date_info.getTime() + tzOffset);
    
    const yyyy = local_date.getFullYear();
    const mm = String(local_date.getMonth() + 1).padStart(2, '0');
    const dd = String(local_date.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  const s = String(val).trim();
  if (s.length === 6 && /^\d+$/.test(s)) {
    // "260714" -> "2026-07-14"
    return `20${s.substring(0, 2)}-${s.substring(2, 4)}-${s.substring(4, 6)}`;
  }
  return s;
}

function parseExcel(triggerPush = false) {
  if (!fs.existsSync(EXCEL_FILE)) {
    console.error('오류: 엑셀 파일이 존재하지 않습니다.');
    return;
  }

  try {
    const workbook = XLSX.readFile(EXCEL_FILE);
    const academyData = {
      wifi: { ssid: "okestro_guest", pw: "Okguest!@" },
      attendance: { url: "", qr: "" },
      solutions: [],
      vpnAccounts: [],
      downloads: [],
      practiceMaterials: [],
      tromboneServices: [],
      middlewares: []
    };

    // 1. 기본정보 시트 파싱
    const infoSheet = workbook.Sheets['기본정보'];
    if (infoSheet) {
      const rows = XLSX.utils.sheet_to_json(infoSheet, { header: 1, defval: "" });
      
      // 와이파이 정보 (1행 SSID, 2행 PW)
      if (rows[1] && rows[1][1]) academyData.wifi.ssid = String(rows[1][1]).trim();
      if (rows[2] && rows[2][1]) academyData.wifi.pw = String(rows[2][1]).trim();
      
      // 출석체크 링크 & QR 코드 검색 (A열 키워드 매칭)
      rows.forEach(row => {
        const key = String(row[0] || "").trim().toUpperCase();
        if (key.includes("URL") || key.includes("출석체크 바로가기")) {
          academyData.attendance.url = String(row[1] || "").trim();
        }
        if (key.includes("QR") || key.includes("코드")) {
          academyData.attendance.qr = String(row[1] || "").trim();
        }
      });
    }

    // 2. 솔루션 접속 정보 시트 파싱
    const solSheet = workbook.Sheets['솔루션 접속 정보'];
    const adminSheet = workbook.Sheets['(관리자)솔루션 접속 정보'];
    const hostMap = {};

    if (adminSheet) {
      const adminRows = XLSX.utils.sheet_to_json(adminSheet, { header: 1, defval: "" });
      adminRows.forEach((row, idx) => {
        if (idx < 1) return;
        const name = String(row[0] || "").trim();
        const host = String(row[3] || "").trim();
        if (name !== "" && host !== "") {
          hostMap[name] = host;
        }
      });
    }

    if (solSheet) {
      const rows = XLSX.utils.sheet_to_json(solSheet, { header: 1, defval: "" });
      let currentSolName = "";
      let currentPw = "";
      let currentUrl = "";
      
      rows.forEach((row, idx) => {
        if (idx < 1) return; // 헤더 스킵
        const solName = String(row[1] || "").trim(); // B열
        const laptopNoText = String(row[2] || "").trim(); // C열 (노트북 번호)
        const studentId = String(row[3] || "").trim(); // D열 (교육생 ID)
        const studentPw = String(row[4] || "").trim(); // E열 (교육생 PW)
        const url = String(row[5] || "").trim(); // F열 (접속주소)
        
        if (solName !== "") {
          currentSolName = solName;
        }
        if (studentPw !== "") {
          currentPw = studentPw;
        }
        if (url !== "") {
          currentUrl = url;
        }
        
        if (currentSolName !== "") {
          // 노트북 번호 정수 추출 (예: "노트북 5번" -> 5)
          let laptopNo = null;
          if (laptopNoText !== "" && laptopNoText !== "-") {
            const numMatch = laptopNoText.match(/\d+/);
            if (numMatch) {
              laptopNo = parseInt(numMatch[0], 10);
            }
          }

          // 관리자 시트에서 매칭되는 host 검색 (부분 매칭 포함)
          let matchedHost = hostMap[currentSolName] || "";
          if (!matchedHost) {
            // "OKESTRO CMP 3.0.5\n(관리자)" 등 형태에 대비해 부분 매칭 시도
            const cleanKey = currentSolName.replace(/\s+/g, "").toLowerCase();
            const matchedKey = Object.keys(hostMap).find(k => {
              const cleanK = k.replace(/\s+/g, "").toLowerCase();
              return cleanKey.includes(cleanK) || cleanK.includes(cleanKey);
            });
            if (matchedKey) {
              matchedHost = hostMap[matchedKey];
            }
          }

          academyData.solutions.push({
            name: currentSolName,
            no: laptopNo,
            id: studentId,
            pw: currentPw,
            url: currentUrl,
            host: matchedHost
          });
        }
      });
    }

    // 3. 실습용 VPN 계정 정보 시트 파싱
    const vpnSheet = workbook.Sheets['실습용 VPN 계정 정보'];
    if (vpnSheet) {
      const rows = XLSX.utils.sheet_to_json(vpnSheet, { header: 1, defval: "" });
      rows.forEach((row, idx) => {
        if (idx < 3) return; // 헤더 설명 및 컬럼명 스킵
        const no = parseInt(row[0], 10);
        const id = String(row[1] || "").trim();
        const pw = String(row[2] || "").trim();
        if (id !== "" && !isNaN(no)) {
          academyData.vpnAccounts.push({ no, id, pw });
        }
      });
    }

    // 4. 교육자료 다운로드 시트 파싱
    const dlSheet = workbook.Sheets['교육자료 다운로드'];
    if (dlSheet) {
      const rows = XLSX.utils.sheet_to_json(dlSheet, { header: 1, defval: "" });
      rows.forEach((row, idx) => {
        if (idx < 1) return;
        const rawDate = row[0];
        const name = String(row[1] || "").trim();
        const url = String(row[2] || "").trim();
        
        if (name !== "") {
          let dateStr = "";
          if (rawDate !== undefined && rawDate !== null && rawDate !== "") {
            dateStr = formatExcelDate(rawDate);
          }
          academyData.downloads.push({ date: dateStr, name, url });
        }
      });
    }

    // 4-2. 실습자료 다운로드 시트 파싱
    const pmSheet = workbook.Sheets['실습자료 다운로드'];
    if (pmSheet) {
      const rows = XLSX.utils.sheet_to_json(pmSheet, { header: 1, defval: "" });
      rows.forEach((row, idx) => {
        if (idx < 1) return;
        const name = String(row[0] || "").trim();
        const url = String(row[1] || "").trim();
        if (name !== "") {
          academyData.practiceMaterials.push({ name, url });
        }
      });
    }

    // 5. Trombone 미들웨어 접속정보 시트 파싱
    const mwSheet = workbook.Sheets['Trombone 미들웨어 접속정보'];
    if (mwSheet) {
      const rows = XLSX.utils.sheet_to_json(mwSheet, { header: 1, defval: "" });
      rows.forEach((row, idx) => {
        if (idx < 1) return;
        const name = String(row[0] || "").trim();
        const id = String(row[1] || "").trim();
        const pw = String(row[2] || "").trim();
        if (name !== "") {
          academyData.middlewares.push({ name, id, pw });
        }
      });
    }

    // 6. TROMBONE 서비스 접속 정보 시트 파싱
    const tbServiceSheet = workbook.Sheets['TROMBONE 서비스 접속 정보'];
    if (tbServiceSheet) {
      const rows = XLSX.utils.sheet_to_json(tbServiceSheet, { header: 1, defval: "" });
      for (let idx = 1; idx < rows.length; idx++) {
        const row = rows[idx];
        const vpnId = String(row[0] || "").trim();
        const userId = String(row[1] || "").trim();
        const bizCode = String(row[2] || "").trim();
        const gitRepo = String(row[3] || "").trim();
        const env = String(row[4] || "").trim();
        const url = String(row[5] || "").trim();
        
        if (vpnId.startsWith("handson-") || vpnId === "강사") {
          let no = 0;
          if (vpnId === "강사") {
            no = 0;
          } else {
            no = parseInt(vpnId.replace("handson-", ""), 10);
          }
          
          const nextRow = rows[idx + 1] || [];
          const nextEnv = String(nextRow[4] || "").trim();
          const nextUrl = String(nextRow[5] || "").trim();
          
          let stgUrl = "";
          let prdUrl = "";
          if (env === "STG") stgUrl = url;
          if (env === "PRD") prdUrl = url;
          if (nextEnv === "STG") stgUrl = nextUrl;
          if (nextEnv === "PRD") prdUrl = nextUrl;

          academyData.tromboneServices.push({
            no,
            vpnId,
            userId,
            bizCode,
            gitRepo,
            stgUrl,
            prdUrl
          });
          
          idx++; // PRD 행 스킵
        }
      }
    }

    // data.js 파일 생성
    const jsContent = `// Automatically generated from Excel file
const academyData = ${JSON.stringify(academyData, null, 2)};

window.ACADEMY_DATA = academyData;

if (typeof module !== 'undefined' && module.exports) {
  module.exports = academyData;
}`;
    fs.writeFileSync(OUTPUT_FILE, jsContent, 'utf8');
    console.log(`[${new Date().toLocaleTimeString()}] ✓ data.js 업데이트 완료 (솔루션: ${academyData.solutions.length}개, 미들웨어: ${academyData.middlewares.length}개, 다운로드: ${academyData.downloads.length}개)`);

    // 변경 사항이 있을 때 깃허브 및 Vercel 실시간 배포 동기화 실행
    if (triggerPush) {
      autoGitPush();
    }

  } catch (err) {
    console.error('엑셀 파싱 중 에러 발생:', err.message);
  }
}

let isPushing = false;
function autoGitPush() {
  if (isPushing) {
    console.log('⚠️ 현재 GitHub 동기화가 진행 중입니다. 대기해 주세요.');
    return;
  }
  isPushing = true;
  console.log('🚀 [GitHub & Vercel] 실시간 배포 동기화(Git Push)를 진행합니다...');
  
  const { exec } = require('child_process');
  exec('git add data.js && (git diff-index --quiet HEAD -- data.js || git commit -m "Auto-update data.js from Excel" && git push)', (err, stdout, stderr) => {
    isPushing = false;
    if (err) {
      console.error('❌ Git push 동기화 실패:', err.message);
      return;
    }
    console.log('✓ [GitHub & Vercel] 실시간 동기화 완료! (Vercel이 백그라운드 배포를 진행 중입니다.)');
  });
}

// 최초 1회 실행 (푸시 없음)
parseExcel(false);

// 엑셀 파일 실시간 감시 시작 (숨김 파일 및 엑셀 임시파일 ~$ 무시)
const watcher = chokidar.watch(EXCEL_FILE, {
  persistent: true,
  ignoreInitial: true,
  ignored: [
    /(^|[\/\\])\../,
    /~\$/
  ]
});

watcher.on('change', () => {
  console.log('엑셀 파일 변경이 감지되었습니다. 재파싱 및 실시간 배포를 진행합니다...');
  setTimeout(() => parseExcel(true), 500);
});

watcher.on('error', error => console.error('감시 오류 발생:', error));

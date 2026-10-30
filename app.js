// app.js
const LIFF_ID = "2011084676-suwQE7jT"; 
const GAS_URL = "https://script.google.com/macros/s/AKfycbzdFBExPsp296HKlcZEK4NZ_qCIfBDbV5REPEYqMyJ4h93vZznQv-yki4XVMMo6Dj-7XQ/exec";

const timeSlots = [
  "08:00-09:00", "09:00-10:00", "10:00-11:00", "11:00-12:00",
  "12:00-13:30", "13:30-14:30", "14:30-15:30", "15:30-16:30",
  "16:30-17:30", "18:00-19:00", "19:00-20:00", "20:00-21:00", "21:00-22:00"
];

let currentUser = null;
let systemData = null;

window.onload = async () => {
  try {
    // 檢查是否有填寫 LIFF ID
    if (LIFF_ID === "YOUR_LIFF_ID_HERE" || !LIFF_ID) {
      alert("錯誤：請先在 app.js 中填寫真實的 LIFF_ID！");
      return;
    }

    await liff.init({ liffId: LIFF_ID });
    
    if (!liff.isLoggedIn()) {
      liff.login();
      return;
    }
    
    const profile = await liff.getProfile();
    fetchInitData(profile.userId);
    
  } catch (err) {
    console.error("LIFF 錯誤詳情:", err);
    // 這裡會把真正的錯誤原因印出來！
    alert("初始化失敗！\n代碼：" + err.code + "\n訊息：" + err.message);
  }
};


async function apiCall(action, payload) {
  document.getElementById('loadingModal').style.display = 'flex';
  const res = await fetch(GAS_URL, {
    method: 'POST',
    body: JSON.stringify({ action, payload })
  });
  const data = await res.json();
  document.getElementById('loadingModal').style.display = 'none';
  if (!data.success) throw new Error(data.error);
  return data.data;
}

async function fetchInitData(userId) {
  try {
    systemData = await apiCall('getInitData', { userId });
    currentUser = { userId: userId };
    
    if (!systemData.isRegistered) {
      document.getElementById('registerModal').style.display = 'flex';
    } else {
      currentUser = systemData.userData;
      setupDashboard();
    }
  } catch (err) {
    alert("載入資料失敗：" + err.message);
  }
}

async function register() {
  const name = document.getElementById('regName').value.trim();
  const nickname = document.getElementById('regNick').value.trim();
  if (!name || !nickname) return alert("請填寫完整資訊");

  try {
    await apiCall('register', { userId: currentUser.userId, name, nickname });
    document.getElementById('registerModal').style.display = 'none';
    fetchInitData(currentUser.userId); // 重新載入
  } catch (err) {
    alert("註冊失敗");
  }
}

function setupDashboard() {
  document.getElementById('userInfoCard').style.display = 'block';
  document.getElementById('scheduleCard').style.display = 'block';
  
  document.getElementById('userName').innerText = currentUser.nickname;
  document.getElementById('roomPwd').innerText = systemData.password;
  
  if (currentUser.isAdmin) {
    document.getElementById('adminBtn').classList.remove('hidden');
  }

  // 設定跑馬燈與彈窗公告
  if (systemData.announcements.length > 0) {
    const annText = systemData.announcements.join(" ｜ ");
    document.getElementById('marqueeText').innerText = annText;
    document.getElementById('marqueeDiv').classList.remove('hidden');
    
    // 隨機彈出公告
    const randomAnn = systemData.announcements[Math.floor(Math.random() * systemData.announcements.length)];
    document.getElementById('popupText').innerText = randomAnn;
    document.getElementById('popupModal').style.display = 'flex';
  }

  renderTimetable();
}

function closePopup() {
  document.getElementById('popupModal').style.display = 'none';
}

function renderTimetable() {
  const container = document.getElementById('timetable');
  container.innerHTML = '';
  
  const days = ['時間', '一', '二', '三', '四', '五'];
  days.forEach(d => {
    const div = document.createElement('div');
    div.className = 'cell header-cell';
    div.innerText = d;
    container.appendChild(div);
  });

  const now = new Date();
  const currentDay = now.getDay(); // 1=Mon, 5=Fri
  const currentHourText = now.getHours() + ":" + String(now.getMinutes()).padStart(2, '0');

  timeSlots.forEach(slot => {
    // 時間欄
    const timeDiv = document.createElement('div');
    timeDiv.className = 'cell time-cell';
    timeDiv.innerText = slot.replace("-", "\n|");
    container.appendChild(timeDiv);

    // 星期一到五
    for (let day = 1; day <= 5; day++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      
      const slotEndHour = slot.split('-')[1]; // ex: "09:00"
      const isExpired = (day < currentDay && currentDay >=1 && currentDay <=5) || (day === currentDay && currentHourText > slotEndHour);

      // 尋找此格子的預約狀態
      const booking = systemData.schedule.find(s => s.day == day && s.timeSlot == slot);

      if (isExpired) {
        cell.classList.add('expired');
        cell.innerText = booking ? (booking.type==='fixed' ? booking.label : booking.nickname) : '';
      } else if (booking) {
        if (booking.type === 'fixed') {
          cell.classList.add('fixed');
          cell.innerText = booking.label;
        } else if (booking.userId === currentUser.userId) {
          cell.classList.add('booked', 'mine');
          cell.innerText = '取消預約';
          cell.onclick = () => handleCancel(day, slot);
        } else {
          cell.classList.add('booked');
          cell.innerText = booking.nickname;
        }
      } else {
        cell.classList.add('available');
        cell.innerText = '可預約';
        cell.onclick = () => handleBook(day, slot);
      }
      container.appendChild(cell);
    }
  });
}

async function handleBook(day, timeSlot) {
  if(!confirm(`確定要預約 星期${['一','二','三','四','五'][day-1]} ${timeSlot} 嗎？`)) return;
  try {
    await apiCall('book', { userId: currentUser.userId, nickname: currentUser.nickname, day, timeSlot });
    alert("預約成功！");
    fetchInitData(currentUser.userId);
  } catch (err) {
    alert(err.message);
  }
}

async function handleCancel(day, timeSlot) {
  if(!confirm(`確定要取消 星期${['一','二','三','四','五'][day-1]} ${timeSlot} 的預約嗎？`)) return;
  try {
    await apiCall('cancel', { userId: currentUser.userId, day, timeSlot });
    alert("取消成功！");
    fetchInitData(currentUser.userId);
  } catch (err) {
    alert(err.message);
  }
}

function openAdmin() {
  const newPwd = prompt("輸入新的社辦密碼 (取消則不修改):");
  if(newPwd) {
    apiCall('adminUpdate', { userId: currentUser.userId, type: 'password', data: { newPassword: newPwd }})
      .then(() => { alert("密碼更新成功！"); fetchInitData(currentUser.userId); });
  }
}

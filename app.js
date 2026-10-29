const LIFF_ID = '2011084676-suwQE7jT'; 
const GAS_API_URL = 'https://script.google.com/macros/s/AKfycby06baHoa-Hg0d6x5NHwGNmtW0ubsK6hPtkAZA0w4RDkmtpd6FTj6rJXUlK89I6_B4U9w/exec';

const timeSlots = [
  "08:00 - 09:00", "09:00 - 10:00", "10:00 - 11:00", "11:00 - 12:00",
  "12:00 - 13:30", "13:30 - 14:30", "14:30 - 15:30", "15:30 - 16:30",
  "16:30 - 17:30", "18:00 - 19:00", "19:00 - 20:00", "20:00 - 21:00", "21:00 - 22:00"
];

let state = {
  user: null,
  bookings: [],
  announcements: [],
  isAdminMode: false
};

async function main() {
  await liff.init({ liffId: LIFF_ID });
  if (!liff.isLoggedIn()) {
    liff.login();
    return;
  }
  
  const profile = await liff.getProfile();
  fetchData('init', { userId: profile.userId });
}

async function fetchData(action, payload) {
  try {
    const response = await fetch(GAS_API_URL, {
      method: 'POST',
      body: JSON.stringify({ action, ...payload })
    });
    const data = await response.json();
    handleApiResponse(action, data, payload.userId);
  } catch (err) {
    alert("系統連線錯誤，請稍後再試！");
  }
}

function handleApiResponse(action, data, currentUserId) {
  if (action === 'init') {
    if (!data.registered) {
      document.getElementById('registerModal').style.display = 'flex';
      return;
    }
    // 已註冊
    state.user = data.user;
    state.bookings = data.bookings;
    state.announcements = data.announcements;
    
    document.getElementById('userInfo').innerText = `你好，${state.user.nickname}`;
    document.getElementById('passwordArea').style.display = 'block';
    document.getElementById('clubPassword').innerText = data.password;

    if (state.user.role === 'admin') {
      document.getElementById('adminBtn').style.display = 'block';
      document.getElementById('adminBtn').onclick = () => {
        state.isAdminMode = true;
        document.getElementById('adminModal').style.display = 'flex';
      };
    }

    renderMarqueeAndModal();
    renderTable();
  } else if (data.success) {
    if (data.bookings) state.bookings = data.bookings;
    if (data.announcements) state.announcements = data.announcements;
    if (data.password) document.getElementById('clubPassword').innerText = data.password;
    renderTable();
    if (action !== 'adminSetFixed') alert('操作成功！');
  } else {
    alert('操作失敗：' + data.error);
  }
}

function renderMarqueeAndModal() {
  if (state.announcements.length > 0) {
    const texts = state.announcements.map(a => a.text).join(' &nbsp;&nbsp; | &nbsp;&nbsp; ');
    document.getElementById('marqueeText').innerHTML = texts;
    
    // 隨機彈出公告
    const randomOne = state.announcements[Math.floor(Math.random() * state.announcements.length)];
    document.getElementById('randomAnnouncementText').innerText = randomOne.text;
    document.getElementById('announcementModal').style.display = 'flex';
  } else {
    document.getElementById('marqueeText').innerText = '目前無最新公告';
  }
}

function renderTable() {
  const tbody = document.getElementById('timetableBody');
  tbody.innerHTML = '';
  
  const now = new Date();
  const currentDay = now.getDay(); // 0(Sun) - 6(Sat)
  const currentHour = now.getHours();
  const isWeekend = (currentDay === 0 || currentDay === 6);

  timeSlots.forEach((slotLabel, timeIndex) => {
    let tr = document.createElement('tr');
    let tdTime = document.createElement('td');
    tdTime.innerText = slotLabel;
    tr.appendChild(tdTime);

    // slot 小時起始值 (簡單切割字串)
    const slotStartHour = parseInt(slotLabel.split(':')[0]);

    for (let day = 1; day <= 5; day++) {
      let td = document.createElement('td');
      td.classList.add('cell-slot');
      
      // 找出該時段的預約紀錄
      const booking = state.bookings.find(b => b.day == day && b.timeIndex == timeIndex);
      
      // 判斷是否過期 (非週末，且該天已過，或同天但小時已過)
      const isExpired = !isWeekend && (day < currentDay || (day === currentDay && currentHour >= slotStartHour));

      if (isExpired && (!booking || booking.type !== 'fixed')) {
        td.classList.add('status-expired');
        td.innerText = booking ? booking.nickname : '';
        td.onclick = () => { if(state.isAdminMode) adminSlotClick(day, timeIndex); else alert('該時段已過期'); };
      } else if (booking) {
        if (booking.type === 'fixed') {
          td.classList.add('status-fixed');
          td.innerText = booking.label;
        } else {
          td.classList.add('status-booked');
          td.innerText = booking.nickname;
          // 若是自己預約的，可以取消
          if (booking.userId === state.user.userId && !isExpired) {
            td.innerText += "\n(點擊取消)";
            td.onclick = () => {
              if (confirm('確定要取消這個時段的預約嗎？')) fetchData('cancel', { userId: state.user.userId, day, timeIndex });
            };
          } else if (state.isAdminMode) {
             td.onclick = () => adminSlotClick(day, timeIndex);
          }
        }
      } else {
        td.classList.add('status-available');
        td.innerText = '可預約';
        td.onclick = () => {
          if (state.isAdminMode) adminSlotClick(day, timeIndex);
          else if (confirm(`確認預約 星期${["一","二","三","四","五"][day-1]} ${slotLabel} 嗎？`)) {
            fetchData('book', { userId: state.user.userId, day, timeIndex });
          }
        };
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
}

// 註冊
async function submitRegister() {
  const name = document.getElementById('regName').value.trim();
  const nickname = document.getElementById('regNickname').value.trim();
  if (!name || !nickname) return alert('請填寫完整資訊');
  
  const profile = await liff.getProfile();
  document.getElementById('registerModal').style.display = 'none';
  fetchData('register', { userId: profile.userId, name, nickname });
}

// 關閉 Modal
function closeModal(id) {
  document.getElementById(id).style.display = 'none';
  if (id === 'adminModal') state.isAdminMode = false;
}

// --- 管理員操作 ---
function adminSlotClick(day, timeIndex) {
  const label = prompt('【管理員設定固定時段】\n請輸入標籤(例如: 樂器班)。\n若要清除該時段設定，請留空並確認：');
  if (label !== null) {
    fetchData('adminSetFixed', { userId: state.user.userId, day, timeIndex, label });
  }
}

function adminUpdatePassword() {
  const pwd = document.getElementById('newPassword').value.trim();
  if(!pwd) return alert('不可為空');
  fetchData('adminUpdatePassword', { userId: state.user.userId, password: pwd });
}

function adminAddAnnouncement() {
  const text = document.getElementById('newAnnouncement').value.trim();
  if(!text) return alert('不可為空');
  fetchData('adminAddAnnouncement', { userId: state.user.userId, text });
}

main();

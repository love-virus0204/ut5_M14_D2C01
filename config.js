/* ===== API_URL ===== */
window.api_url = 'https://script.google.com/macros/s/AKfycbz5TXqNou2gdKn1JXEiqTvAGjpwnT0BucbI3HCBkZfgwHH922uvRIwavS8WZYC_6soQ/exec';

/* ===== min Tools ===== */
window.epoch = Date.UTC(1899,11,30);
const sbT = new Intl.DateTimeFormat("zh-TW", {
    timeZone: "Asia/Taipei",
    hour12: false,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });
window.now = "";
function tick() { now = sbT.format(new Date()); }
tick();
setInterval(tick, 1000);

window.mins = 0;
function tick2() { mins = Number(now.slice(11, 13)) * 60 + Number(now.slice(14, 16)); }
tick2();
setInterval(tick2, 60000);

// 依 序號 轉 yyyy/mn/dd ✔
window.serialToYMD = function(serial) {
  const date = new Date(epoch + (serial * 86400000));
  const yyyy = date.getFullYear();
  const m2 = String(date.getMonth() + 1).padStart(2, "0");
  const d2 = String(date.getDate()).padStart(2, "0");
  return `${yyyy}/${m2}/${d2}`;
}

// 依 yyyy/mm/dd 取日期序號
window.ymdToSerial = function(Ntime){
  const ymd = Ntime.slice(0, 10);
  const [y,m,d] = ymd.split('/').map(Number);
  return ((Date.UTC(y, m-1, d) - epoch) / 86400000);
}

// 依 序號 轉 mn/dd ✔
window.serialToMD = function(sexsb) {
  const date = new Date(epoch + (sexsb * 86400000));
  const y4 = date.getFullYear();
  const m2 = String(date.getMonth() + 1).padStart(2, "0");
  const d2 = String(date.getDate()).padStart(2, "0");
  return `${m2}/${d2}`;
}

// yyyy/mm/dd → yyyymm ✔
window.ymdToYM = function(Ntime){
  //const ymd = Ntime.slice(0, 10);
  const [y, m, d] = Ntime.split('/');
  return Number(y + m.padStart(2,'0'));
}


/* ===== 音效清單 ===== */
window.SOUND_LIST = [
  './mus/levelup.mp3',
  './mus/clock01.wav',
  './mus/clock02.mp3',
  './mus/clock03.mp3'
];

/* ===== 預載所有音效 ===== */
window.SFX_POOL = window.SOUND_LIST.map(src => {
  const a = new Audio(src);
  a.preload = 'auto';
  a.load();
  return a;
});

/* ===== 播放指定索引 ===== */
window.playSFX = function(i){
  const a = window.SFX_POOL[i];
  if(!a) return;
  try {
    a.pause();
    a.currentTime = 0;
    a.volume = 1.0;
    a.loop = false;
    a.play();
  } catch(e){}
};

window.events = ['pointerdown', 'mousedown', 'touchstart', 'keydown', 'wheel'];

window.events = ['pointerdown','pointerup', 'mousedown', 'mouseup', 'touchstart', 'touchend', 'keydown', 'keyup', 'wheel', 'scroll', 'click', 'dblclick','contextmenu'
];

window.onInteract = () => {
  const bgm = document.getElementById('bgm');
  if (!bgm) return;
  window.bgm = bgm;
  window.playing = true;
  bgm.onended = () => (window.playing = true);
};

window.playBGM = () => {
  if (!window.playing || !window.bgm) return;
  window.playing = false;
  window.bgm.volume = 0.7;
  window.bgm.play().catch(() => (window.playing = true));
};

for (const ev of window.events) {
  window.addEventListener(ev, window.onInteract, { once: true, passive: true });
  window.addEventListener(ev, window.playBGM, { passive: true });
}

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('./sw.js')
    .then(r => console.log('SW registered', r.scope))
    .catch(console.error);
}
// 테마 전환: 명시 선택은 data-theme으로, 미선택 시 시스템 설정을 따른다
const root = document.documentElement;
const themeToggle = document.getElementById('themeToggle');
const themeLabel = document.getElementById('themeLabel');
const systemDark = window.matchMedia('(prefers-color-scheme: dark)');

function currentTheme() {
  const explicit = root.getAttribute('data-theme');
  if (explicit === 'dark' || explicit === 'light') return explicit;
  return systemDark.matches ? 'dark' : 'light';
}
function renderThemeLabel() {
  themeLabel.textContent = currentTheme() === 'dark' ? '라이트 모드' : '다크 모드';
}
try {
  const saved = localStorage.getItem('theme');
  if (saved === 'dark' || saved === 'light') root.setAttribute('data-theme', saved);
} catch (e) { /* 저장소 사용 불가 */ }
renderThemeLabel();
systemDark.addEventListener('change', renderThemeLabel);

themeToggle.addEventListener('click', () => {
  const next = currentTheme() === 'dark' ? 'light' : 'dark';
  root.setAttribute('data-theme', next);
  try { localStorage.setItem('theme', next); } catch (e) { /* 저장소 사용 불가 */ }
  renderThemeLabel();
});

// 현재 보고 있는 섹션을 목차에 표시
const tocLinks = Array.from(document.querySelectorAll('.toc a'));
const sections = tocLinks
  .map(a => document.querySelector(a.getAttribute('href')))
  .filter(Boolean);

function setActive(id) {
  tocLinks.forEach(a => a.classList.toggle('active', a.getAttribute('href') === '#' + id));
}
if ('IntersectionObserver' in window && sections.length) {
  const observer = new IntersectionObserver((entries) => {
    const visible = entries.filter(e => e.isIntersecting)
      .sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top);
    if (visible.length) setActive(visible[0].target.id);
  }, { rootMargin: '-20% 0px -65% 0px', threshold: 0 });
  sections.forEach(s => observer.observe(s));
  setActive(sections[0].id);
}

// 이메일 복사
const copyBtn = document.getElementById('copyEmail');
const emailText = document.getElementById('emailText');
copyBtn.addEventListener('click', () => {
  const done = () => {
    copyBtn.textContent = '복사됨';
    copyBtn.classList.add('done');
    setTimeout(() => { copyBtn.textContent = '복사'; copyBtn.classList.remove('done'); }, 1600);
  };
  const fallback = () => {
    const range = document.createRange();
    range.selectNodeContents(emailText);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(emailText.textContent).then(done).catch(fallback);
  } else {
    fallback();
  }
});

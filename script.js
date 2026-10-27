// ===== Theme Toggle =====
const themeToggle = document.getElementById('themeToggle');
const themeIcon = themeToggle.querySelector('.theme-icon');
const html = document.documentElement;

function setTheme(theme, persist) {
  html.setAttribute('data-theme', theme);
  themeIcon.textContent = theme === 'dark' ? '☀' : '☾';
  if (persist) {
    try { localStorage.setItem('theme', theme); } catch (e) { /* storage unavailable */ }
  }
}

let savedTheme = null;
try { savedTheme = localStorage.getItem('theme'); } catch (e) { /* storage unavailable */ }
if (savedTheme === 'dark' || savedTheme === 'light') {
  setTheme(savedTheme, false);
} else if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
  setTheme('dark', false);
}

themeToggle.addEventListener('click', () => {
  const current = html.getAttribute('data-theme');
  setTheme(current === 'dark' ? 'light' : 'dark', true);
});

// ===== Mobile Menu =====
const mobileMenuBtn = document.getElementById('mobileMenuBtn');
const navLinks = document.querySelector('.nav-links');

function closeMenu() {
  navLinks.classList.remove('open');
  mobileMenuBtn.classList.remove('open');
  mobileMenuBtn.setAttribute('aria-expanded', 'false');
}

mobileMenuBtn.addEventListener('click', () => {
  const isOpen = navLinks.classList.toggle('open');
  mobileMenuBtn.classList.toggle('open', isOpen);
  mobileMenuBtn.setAttribute('aria-expanded', String(isOpen));
});

navLinks.querySelectorAll('a').forEach(link => {
  link.addEventListener('click', closeMenu);
});

// ===== Nav shadow on scroll =====
const nav = document.querySelector('.nav');
function updateNav() {
  nav.classList.toggle('scrolled', window.scrollY > 10);
}
window.addEventListener('scroll', updateNav, { passive: true });
updateNav();

// ===== Fade-in on scroll =====
const fadeElements = document.querySelectorAll('.fade-in');

if ('IntersectionObserver' in window) {
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });
  fadeElements.forEach(el => observer.observe(el));
} else {
  fadeElements.forEach(el => el.classList.add('visible'));
}

// Stagger cards within each grid
['.area-grid', '.skill-grid', '.case-list', '.timeline'].forEach(selector => {
  document.querySelectorAll(selector + ' .fade-in').forEach((el, i) => {
    el.style.transitionDelay = (i * 0.07) + 's';
  });
});

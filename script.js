// ===== Theme Toggle =====
const themeToggle = document.getElementById('themeToggle');
const themeIcon = themeToggle.querySelector('.theme-icon');
const html = document.documentElement;

function setTheme(theme) {
  html.setAttribute('data-theme', theme);
  themeIcon.textContent = theme === 'dark' ? '\u2600' : '\u263E';
  localStorage.setItem('theme', theme);
}

// Load saved theme or respect system preference
const savedTheme = localStorage.getItem('theme');
if (savedTheme) {
  setTheme(savedTheme);
} else if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
  setTheme('dark');
}

themeToggle.addEventListener('click', () => {
  const current = html.getAttribute('data-theme');
  setTheme(current === 'dark' ? 'light' : 'dark');
});

// ===== Mobile Menu =====
const mobileMenuBtn = document.getElementById('mobileMenuBtn');
const navLinks = document.querySelector('.nav-links');

mobileMenuBtn.addEventListener('click', () => {
  navLinks.classList.toggle('open');
  mobileMenuBtn.classList.toggle('active');
});

// Close mobile menu on link click
navLinks.querySelectorAll('a').forEach(link => {
  link.addEventListener('click', () => {
    navLinks.classList.remove('open');
    mobileMenuBtn.classList.remove('active');
  });
});

// ===== Scroll Animations (IntersectionObserver) =====
const fadeElements = document.querySelectorAll('.fade-in');

const fadeObserver = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) {
      const delay = entry.target.dataset.delay || 0;
      setTimeout(() => {
        entry.target.classList.add('visible');
      }, delay);
      fadeObserver.unobserve(entry.target);
    }
  });
}, {
  threshold: 0.05,
  rootMargin: '50px 0px 0px 0px'
});

// Immediately reveal elements already in viewport on page load
function revealVisibleElements() {
  fadeElements.forEach(el => {
    const rect = el.getBoundingClientRect();
    if (rect.top < window.innerHeight && rect.bottom > 0) {
      el.classList.add('visible');
      fadeObserver.unobserve(el);
    }
  });
}
window.addEventListener('load', revealVisibleElements);
// Also run after a short delay as fallback
setTimeout(revealVisibleElements, 200);

// Add stagger delays for grouped elements
document.querySelectorAll('.skills-grid .skill-category').forEach((el, i) => {
  el.dataset.delay = i * 100;
});
document.querySelectorAll('.timeline-item').forEach((el, i) => {
  el.dataset.delay = i * 120;
});
document.querySelectorAll('.project-card').forEach((el, i) => {
  el.dataset.delay = i * 80;
});

fadeElements.forEach(el => fadeObserver.observe(el));

// ===== Skill Bar Animation =====
const skillFills = document.querySelectorAll('.skill-fill');

const skillObserver = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      const width = entry.target.dataset.width;
      entry.target.style.width = width + '%';
      skillObserver.unobserve(entry.target);
    }
  });
}, {
  threshold: 0.3
});

skillFills.forEach(fill => skillObserver.observe(fill));

// ===== Project Filter =====
const filterBtns = document.querySelectorAll('.filter-btn');
const projectCards = document.querySelectorAll('.project-card');

filterBtns.forEach(btn => {
  btn.addEventListener('click', () => {
    // Update active button
    filterBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');

    const filter = btn.dataset.filter;

    projectCards.forEach(card => {
      if (filter === 'all' || card.dataset.category === filter) {
        card.classList.remove('hidden');
        card.style.animation = 'fadeUp 0.4s ease forwards';
      } else {
        card.classList.add('hidden');
      }
    });
  });
});

// ===== Smooth scroll for nav links =====
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
  anchor.addEventListener('click', function (e) {
    e.preventDefault();
    const target = document.querySelector(this.getAttribute('href'));
    if (target) {
      const offset = 80; // nav height
      const top = target.getBoundingClientRect().top + window.pageYOffset - offset;
      window.scrollTo({ top, behavior: 'smooth' });
    }
  });
});

// ===== Nav background on scroll =====
const nav = document.querySelector('.nav');
window.addEventListener('scroll', () => {
  if (window.scrollY > 50) {
    nav.style.boxShadow = '0 1px 8px rgba(0,0,0,0.08)';
  } else {
    nav.style.boxShadow = 'none';
  }
});

// CSS animation keyframe for filter
const style = document.createElement('style');
style.textContent = `
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(12px); }
    to { opacity: 1; transform: translateY(0); }
  }
`;
document.head.appendChild(style);

// ===== Excel Tabs =====
const excelTabs = document.querySelectorAll('.excel-tab');
const excelPanels = document.querySelectorAll('.excel-panel');

excelTabs.forEach(tab => {
  tab.addEventListener('click', () => {
    excelTabs.forEach(t => t.classList.remove('active'));
    excelPanels.forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('panel-' + tab.dataset.tab).classList.add('active');
  });
});

// ===== Case Study Modal =====
const caseData = {
  case1: {
    title: 'SHA 주주간계약 검토',
    tag: 'IR/투자',
    sections: [
      { title: 'Background', content: 'GXC의 투자 라운드 진행 과정에서 주주간계약(SHA)의 핵심 조항에 대한 체계적 분석이 필요했습니다. 법률 자문 내용을 실무 관점에서 정리하고, 향후 투자 계약 시 빠르게 참조할 수 있는 가이드를 만드는 것이 목표였습니다.' },
      { title: 'Process', list: ['ROFR(우선매수권) 조항의 실무 적용 시나리오 분석', 'Drag-Along / Tag-Along 권리의 행사 조건 정리', '희석방지(Anti-dilution) 조항의 유형별 비교 (Full Ratchet vs Weighted Average)', '투자자 동의권(Consent Rights) 범위 및 영향 분석', '법률 자문사 미팅 내용 정리 및 실무 Q&A 문서화'] },
      { title: 'Result', result: '5개 핵심 조항에 대한 실무 가이드 문서 완성. 향후 투자 계약 검토 시 체크리스트로 활용 가능한 레퍼런스 구축.' },
      { title: 'Lessons Learned', content: '법률 용어를 실무 관점으로 번역하는 과정에서 투자 구조에 대한 깊은 이해를 얻었습니다. 특히 희석방지 조항이 후속 투자 라운드에 미치는 영향을 구체적으로 파악할 수 있었습니다.' }
    ]
  },
  case2: {
    title: '월말 결산 프로세스 혁신',
    tag: '재무/회계',
    sections: [
      { title: 'Background', content: '매월 결산 시 수작업 위주의 프로세스로 인해 3일 이상 소요되었고, 휴먼 에러 리스크가 높았습니다. 체계적인 결산 프로세스 구축과 Excel 자동화를 통한 효율화가 필요했습니다.' },
      { title: 'Process', list: ['기존 결산 프로세스 전수조사 및 병목 구간 파악', '전표 데이터 정리 자동화 (XLOOKUP + 피벗테이블 활용)', '계정과목별 자동 집계 템플릿 구축', 'VBA 매크로를 활용한 보고서 자동 생성', '결산 체크리스트 및 SOP 문서 작성'] },
      { title: 'Result', result: '결산 소요 시간 3일 → 1.2일로 약 60% 단축. 표준화된 SOP로 인한 담당자 변경 시에도 일관된 결산 품질 유지.' },
      { title: 'Lessons Learned', content: 'Excel 고급 기능만으로도 상당한 업무 자동화가 가능하다는 것을 체감했습니다. 프로세스를 먼저 정리하고 자동화하는 순서가 중요합니다.' }
    ]
  },
  case3: {
    title: 'TTC 기반 보상 체계 설계',
    tag: 'HR/노무',
    sections: [
      { title: 'Background', content: '스타트업 성장에 따라 직급/직무별 보상 기준이 필요했습니다. 시장 데이터 기반의 TTC(Total Target Compensation) 체계를 도입하여 공정하고 경쟁력 있는 보상 구조를 만들고자 했습니다.' },
      { title: 'Process', list: ['국내 IT 스타트업 보상 시장조사 및 벤치마킹', 'TTC 구성요소 정의 (기본급 + 성과급 + 스톡옵션)', '직급/직무별 Pay Band 설계', '스톡옵션 배분 기준 및 베스팅 스케줄 수립', '경영진 보고 및 의사결정 자료 준비'] },
      { title: 'Result', result: '전사 적용 가능한 보상 체계 프레임워크 구축 완료. 채용 시 보상 제안의 일관성과 투명성 확보.' },
      { title: 'Lessons Learned', content: '보상 설계는 재무(비용 관리)와 HR(인재 확보)의 교차점에 있어, 크로스펑셔널 역량이 핵심이라는 것을 깊이 느꼈습니다.' }
    ]
  }
};

function openCaseModal(caseId) {
  const data = caseData[caseId];
  if (!data) return;

  let html = `<h2>${data.title}</h2><span class="case-tag">${data.tag}</span>`;

  data.sections.forEach(sec => {
    html += `<div class="modal-section"><h4>${sec.title}</h4>`;
    if (sec.content) {
      html += `<p>${sec.content}</p>`;
    }
    if (sec.list) {
      html += '<ul>' + sec.list.map(item => `<li>${item}</li>`).join('') + '</ul>';
    }
    if (sec.result) {
      html += `<div class="modal-result-box"><p>${sec.result}</p></div>`;
    }
    html += '</div>';
  });

  document.getElementById('modalBody').innerHTML = html;
  document.getElementById('caseModal').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeCaseModal() {
  document.getElementById('caseModal').classList.remove('open');
  document.body.style.overflow = '';
}

// Close modal on overlay click
document.getElementById('caseModal').addEventListener('click', (e) => {
  if (e.target === e.currentTarget) closeCaseModal();
});

// Close modal on Escape
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') closeCaseModal();
});

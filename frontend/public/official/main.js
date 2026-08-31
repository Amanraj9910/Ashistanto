document.addEventListener('DOMContentLoaded', () => {
  // Sticky navbar
  const nav = document.querySelector('.nav');
  window.addEventListener('scroll', () => {
    nav.classList.toggle('scrolled', window.scrollY > 10);
  }, { passive: true });

  // Mobile menu
  const burger = document.querySelector('.burger');
  const mob = document.querySelector('.mob-menu');
  if (burger && mob) {
    burger.addEventListener('click', () => {
      mob.classList.toggle('open');
      const s = burger.querySelectorAll('span');
      const isOpen = mob.classList.contains('open');
      s[0].style.transform = isOpen ? 'rotate(45deg) translate(4px, 4px)' : 'none';
      s[1].style.opacity = isOpen ? '0' : '1';
      s[2].style.transform = isOpen ? 'rotate(-45deg) translate(4px, -4px)' : 'none';
    });
    mob.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        mob.classList.remove('open');
        burger.querySelectorAll('span').forEach(s => { s.style.transform = 'none'; s.style.opacity = '1'; });
      });
    });
  }

  // Capability tabs
  document.querySelectorAll('.cap-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const t = tab.dataset.tab;
      document.querySelectorAll('.cap-tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.cap-pane').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById(t)?.classList.add('active');
    });
  });

  // Scroll fade-in
  const obs = new IntersectionObserver(entries => {
    entries.forEach(e => {
      if (e.isIntersecting) { e.target.classList.add('visible'); obs.unobserve(e.target); }
    });
  }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });
  document.querySelectorAll('.fi').forEach(el => obs.observe(el));

  // Counter animation
  const cObs = new IntersectionObserver(entries => {
    entries.forEach(e => {
      if (e.isIntersecting) {
        const el = e.target;
        const end = parseFloat(el.dataset.count);
        const suffix = el.dataset.suffix || '';
        const isFloat = el.dataset.count.includes('.');
        const dur = 1800, start = performance.now();
        const animate = now => {
          const p = Math.min((now - start) / dur, 1);
          const eased = 1 - Math.pow(1 - p, 3);
          el.textContent = (isFloat ? (eased * end).toFixed(1) : Math.floor(eased * end)) + suffix;
          if (p < 1) requestAnimationFrame(animate);
        };
        requestAnimationFrame(animate);
        cObs.unobserve(el);
      }
    });
  }, { threshold: 0.5 });
  document.querySelectorAll('[data-count]').forEach(el => cObs.observe(el));

  // Smooth scroll with nav offset
  document.querySelectorAll('a[href^="#"]').forEach(a => {
    a.addEventListener('click', e => {
      e.preventDefault();
      const target = document.querySelector(a.getAttribute('href'));
      if (target) {
        const top = target.getBoundingClientRect().top + window.pageYOffset - 72;
        window.scrollTo({ top, behavior: 'smooth' });
      }
    });
  });

  // Active nav link on scroll
  const sections = document.querySelectorAll('section[id]');
  const navLinks = document.querySelectorAll('.nav-menu a');
  window.addEventListener('scroll', () => {
    const scrollPos = window.scrollY + 100;
    sections.forEach(sec => {
      if (scrollPos >= sec.offsetTop && scrollPos < sec.offsetTop + sec.offsetHeight) {
        const id = sec.getAttribute('id');
        navLinks.forEach(link => {
          link.style.color = link.getAttribute('href') === '#' + id ? '#c8102e' : '';
        });
      }
    });
  }, { passive: true });
});

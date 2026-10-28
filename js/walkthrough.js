// js/walkthrough.js — Automated walkthrough for screen recording
(function() {
  if (!location.search.includes('walkthrough')) return;

  function wait(ms) {
    return new Promise(r => setTimeout(r, ms));
  }

  function smoothScroll(el, target, duration) {
    return new Promise(resolve => {
      const start = el.scrollTop;
      const change = target - start;
      const startTime = performance.now();

      function animate(now) {
        const elapsed = now - startTime;
        const progress = Math.min(elapsed / duration, 1);
        // easeInOutQuad
        const ease = progress < 0.5 
          ? 2 * progress * progress 
          : -1 + (4 - 2 * progress) * progress;
        el.scrollTop = start + change * ease;

        if (progress < 1) {
          requestAnimationFrame(animate);
        } else {
          el.scrollTop = target;
          resolve();
        }
      }
      requestAnimationFrame(animate);
    });
  }

  window.addEventListener('load', async () => {
    console.log('[Walkthrough] Starting automated presentation...');
    await wait(2200);

    // 1. Home Menu Interaction
    console.log('[Walkthrough] Scene 1: Home Menu');
    for (let idx of [0, 1, 2, 3, 0]) {
      if (window.Controller) Controller.select(idx);
      await wait(750);
    }
    await wait(1000);

    // 2. Scene 2: Profil (About & Polaroid)
    console.log('[Walkthrough] Scene 2: Profil');
    if (window.Controller) Controller.goTo('about');
    await wait(2000);
    const aboutScreen = document.getElementById('screen-about');
    if (aboutScreen) {
      await smoothScroll(aboutScreen, 350, 2500);
      await wait(1800);
      await smoothScroll(aboutScreen, 0, 1500);
      await wait(800);
    }

    // 3. Scene 3: Layanan (Services)
    console.log('[Walkthrough] Scene 3: Layanan');
    if (window.Controller) Controller.goTo('projects');
    await wait(2000);
    const projScreen = document.getElementById('screen-projects');
    if (projScreen) {
      await smoothScroll(projScreen, 500, 2800);
      await wait(1800);
      await smoothScroll(projScreen, 950, 2500);
      await wait(1800);
      await smoothScroll(projScreen, 0, 1600);
      await wait(800);
    }

    // 4. Scene 4: Teknologi (Skills)
    console.log('[Walkthrough] Scene 4: Teknologi');
    if (window.Controller) Controller.goTo('skills');
    await wait(2400); // allow skill bars animation to complete
    const skillScreen = document.getElementById('screen-skills');
    if (skillScreen) {
      await smoothScroll(skillScreen, 550, 2500);
      await wait(1800);
      await smoothScroll(skillScreen, 1100, 2500);
      await wait(1800);
      await smoothScroll(skillScreen, 0, 1600);
      await wait(800);
    }

    // 5. Scene 5: Kontak
    console.log('[Walkthrough] Scene 5: Kontak');
    if (window.Controller) Controller.goTo('contact');
    await wait(2000);
    const contactScreen = document.getElementById('screen-contact');
    if (contactScreen) {
      await smoothScroll(contactScreen, 460, 2500);
      await wait(2000);
      await smoothScroll(contactScreen, 0, 1500);
      await wait(800);
    }

    // 6. Scene 6: Return to Home
    console.log('[Walkthrough] Scene 6: Return to Home');
    if (window.Controller) Controller.goTo('home');
    await wait(3500);

    console.log('[Walkthrough] Completed successfully!');
    document.body.setAttribute('data-walkthrough-done', 'true');
  });
})();

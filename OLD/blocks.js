// blocks.js — Blocks configuration & UI logic
(function () {
  function ready(fn) {
    if (document.readyState !== 'loading') fn();
    else document.addEventListener('DOMContentLoaded', fn);
  }

  ready(function () {
    const blocksBtn = document.querySelector('[data-tool="blocks"]');
    const blocksSheet = document.getElementById('blocks-sheet');
    const blocksClose = document.getElementById('blocks-close');
    
    const pvSheet = document.getElementById('pv-sheet');
    const zoneSheet = document.getElementById('zone-sheet');
    const drawer = document.getElementById('drawer');
    const searchWrap = document.getElementById('search-wrap');

    // Helpers
    function closeAllOthers() {
      if (pvSheet) pvSheet.classList.remove('open');
      if (zoneSheet) zoneSheet.classList.remove('open');
      if (drawer) drawer.classList.remove('open');
      if (searchWrap) {
        searchWrap.classList.remove('open');
        searchWrap.setAttribute('aria-expanded', 'false');
      }
    }

    // Toggle Blocks Sheet
    blocksBtn?.addEventListener('click', () => {
      const willOpen = !blocksSheet.classList.contains('open');
      if (willOpen) {
        closeAllOthers();
        blocksBtn.classList.add('is-primary');
      } else {
        blocksBtn.classList.remove('is-primary');
      }
      blocksSheet.classList.toggle('open');
    });

    // Close Button
    blocksClose?.addEventListener('click', () => {
      blocksSheet.classList.remove('open');
      blocksBtn?.classList.remove('is-primary');
    });

    // Close blocks when other tools are opened
    const otherTools = ['pv', 'zoning', 'tools', 'search'];
    otherTools.forEach(tool => {
      const btn = document.querySelector(`[data-tool="${tool}"]`);
      btn?.addEventListener('click', () => {
        blocksSheet.classList.remove('open');
        blocksBtn?.classList.remove('is-primary');
      });
    });

    // Alignment Segment
    const alignSeg = document.getElementById('blk-align');
    alignSeg?.querySelectorAll('button').forEach(btn => {
      btn.addEventListener('click', () => {
        alignSeg.querySelectorAll('button').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        console.log('Block alignment changed to:', btn.dataset.val);
      });
    });

    // Apply Button Placeholder
    const applyBtn = document.getElementById('blk-apply');
    applyBtn?.addEventListener('click', () => {
      console.log('Applying block configuration...');
      // Logic for block modification would go here
      const pattern = document.getElementById('blk-pattern')?.value;
      const minGap = document.getElementById('blk-gap-min')?.value;
      const maxArea = document.getElementById('blk-area-max')?.value;
      
      alert(`Blocks Configuration Updated:\nPattern: ${pattern}\nMin Gap: ${minGap}m\nMax Area: ${maxArea}m²`);
      
      blocksSheet.classList.remove('open');
      blocksBtn?.classList.remove('is-primary');
    });
  });
})();

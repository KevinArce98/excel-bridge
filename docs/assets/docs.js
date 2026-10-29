(() => {
  const copyText = async text => {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      const area = document.createElement('textarea');
      area.value = text;
      area.setAttribute('readonly', '');
      area.style.position = 'fixed';
      area.style.opacity = '0';
      document.body.appendChild(area);
      area.select();
      const done = document.execCommand('copy');
      area.remove();
      return done;
    }
  };

  document.querySelectorAll('.copy').forEach(button => {
    button.addEventListener('click', async () => {
      const code = button.closest('.code').querySelector('code').textContent;
      if (!(await copyText(code))) return;
      const label = button.textContent;
      button.textContent = button.dataset.copied;
      button.dataset.state = 'done';
      setTimeout(() => {
        button.textContent = label;
        delete button.dataset.state;
      }, 1600);
    });
  });

  const groupTabs = () => {
    const blocks = [...document.querySelectorAll('.prose > .code[data-tab]')];
    const sets = [];
    blocks.forEach(block => {
      const last = sets[sets.length - 1];
      if (last && last[last.length - 1].nextElementSibling === block) last.push(block);
      else sets.push([block]);
    });
    sets.forEach((set, setIndex) => {
      if (set.length < 2) return;
      const wrapper = document.createElement('div');
      wrapper.className = 'tabset';
      const list = document.createElement('div');
      list.className = 'tabs';
      list.setAttribute('role', 'tablist');
      set[0].before(wrapper);
      wrapper.append(list);
      const tabs = set.map((block, index) => {
        const tab = document.createElement('button');
        tab.type = 'button';
        tab.setAttribute('role', 'tab');
        tab.id = `tab-${setIndex}-${index}`;
        tab.textContent = block.dataset.tab;
        tab.setAttribute('aria-controls', `panel-${setIndex}-${index}`);
        block.id = `panel-${setIndex}-${index}`;
        block.setAttribute('role', 'tabpanel');
        block.setAttribute('aria-labelledby', tab.id);
        list.append(tab);
        wrapper.append(block);
        return tab;
      });
      const select = index => {
        tabs.forEach((tab, tabIndex) => {
          const selected = tabIndex === index;
          tab.setAttribute('aria-selected', String(selected));
          tab.tabIndex = selected ? 0 : -1;
          set[tabIndex].hidden = !selected;
        });
      };
      tabs.forEach((tab, index) => {
        tab.addEventListener('click', () => select(index));
        tab.addEventListener('keydown', event => {
          const keys = { ArrowRight: 1, ArrowLeft: -1 };
          let target = index;
          if (event.key in keys) target = (index + keys[event.key] + tabs.length) % tabs.length;
          else if (event.key === 'Home') target = 0;
          else if (event.key === 'End') target = tabs.length - 1;
          else return;
          event.preventDefault();
          select(target);
          tabs[target].focus();
        });
      });
      select(0);
    });
  };
  groupTabs();

  const tocLinks = [...document.querySelectorAll('.toc a')];
  if (tocLinks.length && 'IntersectionObserver' in window) {
    const targets = tocLinks.map(link => document.getElementById(link.getAttribute('href').slice(1))).filter(Boolean);
    const visible = new Set();
    const mark = () => {
      const current = targets.find(target => visible.has(target)) ?? null;
      tocLinks.forEach(link => {
        const active = current && link.getAttribute('href') === `#${current.id}`;
        if (active) link.setAttribute('aria-current', 'true');
        else link.removeAttribute('aria-current');
      });
    };
    const observer = new IntersectionObserver(
      entries => {
        entries.forEach(entry => (entry.isIntersecting ? visible.add(entry.target) : visible.delete(entry.target)));
        mark();
      },
      { rootMargin: '-80px 0px -65% 0px' }
    );
    targets.forEach(target => observer.observe(target));
  }

  const fold = document.querySelector('.rail-fold');
  if (fold) {
    const narrow = window.matchMedia('(max-width: 900px)');
    const syncFold = () => {
      fold.open = !narrow.matches;
    };
    syncFold();
    narrow.addEventListener('change', syncFold);
  }
})();

/**
 * A four-card walkthrough shown the first time the app is opened, and
 * repeatable afterwards from the "?" button. Deliberately short: an author who
 * knows Word needs orientation, not a manual.
 */

const SEEN_KEY = 'manuscript-formatter:tour-seen';

interface TourStep {
  title: string;
  body: string;
  /** Inline SVG, so nothing has to load and the page's CSP stays strict. */
  art: string;
}

/** Simple line drawings in the accent colour, sized to the card. */
const ART = {
  choose: `
    <rect x="14" y="16" width="40" height="52" rx="4"/>
    <rect x="66" y="16" width="40" height="52" rx="4"/>
    <path d="M24 30h20M24 40h20M24 50h12"/>
    <path d="M76 30h20M76 40h20M76 50h12"/>
    <circle cx="86" cy="62" r="0.5"/>`,
  check: `
    <rect x="30" y="10" width="60" height="64" rx="4"/>
    <path d="M44 26h32M40 38h40M40 48h40M40 58h24"/>
    <circle cx="86" cy="62" r="14" fill="var(--card)"/>
    <path d="M79 62l5 5 10-11"/>`,
  make: `
    <rect x="20" y="12" width="44" height="58" rx="4"/>
    <path d="M30 28h24M30 38h24M30 48h16"/>
    <path d="M70 41h24M86 33l8 8-8 8"/>
    <rect x="96" y="12" width="8" height="58" rx="2"/>`,
  finish: `
    <rect x="26" y="10" width="56" height="60" rx="4"/>
    <path d="M38 26h32M38 38h32M38 50h20"/>
    <path d="M54 64v14M46 72l8 8 8-8"/>`,
};

const STEPS: TourStep[] = [
  {
    title: '1. Choose a design, add your book',
    body:
      'Pick a book size and a look and the app builds the design for you. Already have a KDP ' +
      'template or a book you like the look of? Switch to “Use my own design” instead.',
    art: ART.choose,
  },
  {
    title: '2. Check the sample pages',
    body:
      'You will see pages built from your own words, at the real size of the printed book. ' +
      'If a chapter title or a scene break looks wrong, you can fix it before anything is made.',
    art: ART.check,
  },
  {
    title: '3. Make the book',
    body:
      'The app writes a brand-new Word file. Your manuscript and your design file are only read, ' +
      'never changed, and nothing is uploaded anywhere.',
    art: ART.make,
  },
  {
    title: '4. Check it, then upload',
    body:
      'Open the new file in Word and flick through it. When you upload to KDP, use Amazon’s own ' +
      'Print Previewer as the final check — it is the one that decides what gets printed.',
    art: ART.finish,
  },
];

/** localStorage is unavailable in some privacy modes; never let that throw. */
function readFlag(key: string): string | null {
  try {
    return window.localStorage.getItem(key);
  } catch {
    return null;
  }
}

function writeFlag(key: string, value: string): void {
  try {
    window.localStorage.setItem(key, value);
  } catch {
    /* Nothing to do: the tour simply shows again next time. */
  }
}

export function hasSeenTour(): boolean {
  return readFlag(SEEN_KEY) === 'yes';
}

export function markTourSeen(seen: boolean): void {
  writeFlag(SEEN_KEY, seen ? 'yes' : 'no');
}

let active: { overlay: HTMLElement; restoreFocus: Element | null } | null = null;

export function startTour(): void {
  if (active) return;
  let index = 0;

  const overlay = document.createElement('div');
  overlay.className = 'tour-overlay';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-labelledby', 'tour-title');

  const card = document.createElement('div');
  card.className = 'tour-card';
  overlay.appendChild(card);

  const close = (remember: boolean): void => {
    if (remember) markTourSeen(true);
    document.removeEventListener('keydown', onKey, true);
    overlay.remove();
    const restore = active?.restoreFocus;
    active = null;
    if (restore instanceof HTMLElement) restore.focus();
  };

  const onKey = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      event.preventDefault();
      close(true);
      return;
    }
    if (event.key !== 'Tab') return;
    // Keep focus inside the dialog while it is open.
    const focusable = card.querySelectorAll<HTMLElement>('button');
    if (focusable.length === 0) return;
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  };

  const draw = (): void => {
    const step = STEPS[index];
    card.replaceChildren();

    const art = document.createElement('div');
    art.className = 'tour-art';
    art.innerHTML =
      `<svg viewBox="0 0 120 88" aria-hidden="true" fill="none" stroke="currentColor" ` +
      `stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${step.art}</svg>`;
    card.appendChild(art);

    const heading = document.createElement('h2');
    heading.id = 'tour-title';
    heading.className = 'tour-title';
    heading.textContent = step.title;
    card.appendChild(heading);

    const body = document.createElement('p');
    body.className = 'tour-body';
    body.textContent = step.body;
    card.appendChild(body);

    const dots = document.createElement('div');
    dots.className = 'tour-dots';
    STEPS.forEach((_, i) => {
      const dot = document.createElement('span');
      dot.className = `tour-dot${i === index ? ' on' : ''}`;
      dots.appendChild(dot);
    });
    card.appendChild(dots);

    const actions = document.createElement('div');
    actions.className = 'tour-actions';

    const skip = document.createElement('button');
    skip.type = 'button';
    skip.className = 'linkish';
    skip.textContent = index === STEPS.length - 1 ? 'Close' : 'Skip';
    skip.addEventListener('click', () => close(true));
    actions.appendChild(skip);

    const spacer = document.createElement('span');
    spacer.className = 'tour-spacer';
    actions.appendChild(spacer);

    if (index > 0) {
      const back = document.createElement('button');
      back.type = 'button';
      back.className = 'secondary';
      back.textContent = 'Back';
      back.addEventListener('click', () => {
        index--;
        draw();
      });
      actions.appendChild(back);
    }

    const next = document.createElement('button');
    next.type = 'button';
    next.className = 'primary';
    next.textContent = index === STEPS.length - 1 ? 'Get started' : 'Next';
    next.addEventListener('click', () => {
      if (index === STEPS.length - 1) close(true);
      else {
        index++;
        draw();
      }
    });
    actions.appendChild(next);

    card.appendChild(actions);
    next.focus();
  };

  overlay.addEventListener('click', (event) => {
    if (event.target === overlay) close(true);
  });

  active = { overlay, restoreFocus: document.activeElement };
  document.body.appendChild(overlay);
  document.addEventListener('keydown', onKey, true);
  draw();
}

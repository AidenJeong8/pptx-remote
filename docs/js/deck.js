// Shared by the presenter and the phone remote: parse a .pptx in the browser,
// render slides into any container, and pull out slide titles and speaker notes.
import { PptxViewer, parseZip, buildPresentation, RECOMMENDED_ZIP_LIMITS } from '../vendor/pptx-renderer.js';

export async function loadDeck(buffer) {
  const pres = buildPresentation(await parseZip(buffer, RECOMMENDED_ZIP_LIMITS));

  // The viewer needs a host element even though we only render into our own containers.
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-100000px;top:0;width:10px;height:10px;overflow:hidden;visibility:hidden';
  document.body.appendChild(host);
  const viewer = new PptxViewer(host, {
    onSlideError: (i, e) => console.warn('slide', i, e),
    onNodeError: (id, e) => console.warn('node', id, e),
  });
  viewer.load(pres);

  // Hidden slides are skipped in the show, as in PowerPoint.
  const order = pres.slides.map((s, i) => (s.hidden ? -1 : i)).filter((i) => i >= 0);
  const { titles, notes } = await readText(buffer, order.map((i) => pres.slides[i]));

  return {
    count: order.length,
    width: viewer.slideWidth,
    height: viewer.slideHeight,
    titles,
    notes,
    // Render visible slide `n` into `el`, scaled to fit `boxW`×`boxH` (or by width only).
    render(n, el, boxW, boxH) {
      const prev = el._slideHandle;
      if (prev) prev.dispose();
      // Reset inherited styles (e.g. from a <button>) so text lays out as on a plain page.
      const box = document.createElement('div');
      box.style.cssText = 'all:initial;display:block';
      el.replaceChildren(box);
      const scale = boxH ? Math.min(boxW / viewer.slideWidth, boxH / viewer.slideHeight) : boxW / viewer.slideWidth;
      const handle = viewer.renderThumbnailToContainer(order[n], box, { scale });
      // The wrapper is white and can peek out past the scaled slide by a subpixel.
      if (box.firstElementChild) box.firstElementChild.style.background = 'transparent';
      el._slideHandle = handle;
      return handle;
    },
    destroy() {
      viewer.destroy?.();
      host.remove();
    },
  };
}

async function readText(buffer, slides) {
  try {
    const zip = await window.JSZip.loadAsync(buffer);
    const xml = (t) => new DOMParser().parseFromString(t, 'application/xml');
    const byName = (doc, name) => [...doc.getElementsByTagName('*')].filter((e) => e.localName === name);
    const text = (el, sep) => byName(el, 'p').map((p) => byName(p, 't').map((t) => t.textContent).join('')).filter(Boolean).join(sep).trim();
    const read = (path) => zip.file(path)?.async('text');

    const results = await Promise.all(slides.map(async (s, n) => {
      const path = s.slidePath || '';
      const slideXml = await read(path);
      const title = (slideXml && findTitle(xml(slideXml), byName, text)) || `Slide ${n + 1}`;

      const rels = await read(path.replace(/([^/]+)$/, '_rels/$1.rels'));
      const rel = rels && byName(xml(rels), 'Relationship').find((r) => /notesSlide$/.test(r.getAttribute('Type')));
      if (!rel) return { title, notes: '' };
      const notesXml = await read(new URL(rel.getAttribute('Target'), 'http://x/' + path).pathname.slice(1));
      if (!notesXml) return { title, notes: '' };
      // Only the body placeholder holds the notes text; skip the slide image and slide number.
      const body = byName(xml(notesXml), 'sp').find((sp) => byName(sp, 'ph').some((ph) => (ph.getAttribute('type') || 'body') === 'body'));
      return { title, notes: body ? text(body, '\n') : '' };
    }));
    return { titles: results.map((r) => r.title), notes: results.map((r) => r.notes) };
  } catch (e) {
    console.warn('slide text', e);
    return { titles: slides.map((_, n) => `Slide ${n + 1}`), notes: slides.map(() => '') };
  }
}

// A title placeholder if there is one; otherwise the topmost text box set at 24pt or larger,
// which skips small eyebrow labels and big stats placed lower on the slide.
function findTitle(doc, byName, text) {
  const shapes = byName(doc, 'sp');
  const ph = shapes.find((sp) => byName(sp, 'ph').some((p) => /^(title|ctrTitle)$/.test(p.getAttribute('type'))));
  if (ph && text(ph, ' ')) return text(ph, ' ');
  let best = null;
  for (const sp of shapes) {
    const t = text(sp, ' ');
    if (!t) continue;
    const maxSz = Math.max(0, ...byName(sp, 'rPr').concat(byName(sp, 'defRPr')).map((r) => Number(r.getAttribute('sz')) || 0));
    if (maxSz < 2400) continue;
    const off = byName(sp, 'off')[0];
    const y = off ? Number(off.getAttribute('y')) || 0 : 0;
    if (!best || y < best.y) best = { y, t };
  }
  return best ? best.t : text(shapes.find((sp) => text(sp, ' ')) || doc, ' ').slice(0, 80);
}

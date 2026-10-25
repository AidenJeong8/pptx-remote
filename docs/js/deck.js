// Shared by the presenter and the phone remote: parse a .pptx in the browser,
// render slides into any container, and pull out speaker notes.
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
  const notes = await readNotes(buffer, order.map((i) => pres.slides[i]));

  return {
    count: order.length,
    width: viewer.slideWidth,
    height: viewer.slideHeight,
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

async function readNotes(buffer, slides) {
  try {
    const zip = await window.JSZip.loadAsync(buffer);
    const xml = (t) => new DOMParser().parseFromString(t, 'application/xml');
    const byName = (doc, name) => [...doc.getElementsByTagName('*')].filter((e) => e.localName === name);
    return await Promise.all(slides.map(async (s) => {
      const path = s.slidePath || '';
      const relsPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
      const rels = await zip.file(relsPath)?.async('text');
      if (!rels) return '';
      const rel = byName(xml(rels), 'Relationship').find((r) => /notesSlide$/.test(r.getAttribute('Type')));
      if (!rel) return '';
      const target = new URL(rel.getAttribute('Target'), 'http://x/' + path).pathname.slice(1);
      const notesXml = await zip.file(target)?.async('text');
      if (!notesXml) return '';
      // Only the body placeholder holds the notes text; skip slide image and slide number.
      const body = byName(xml(notesXml), 'sp').find((sp) => byName(sp, 'ph').some((ph) => (ph.getAttribute('type') || 'body') === 'body' && ph.getAttribute('type') !== 'sldNum'));
      if (!body) return '';
      return byName(body, 'p').map((p) => byName(p, 't').map((t) => t.textContent).join('')).join('\n').trim();
    }));
  } catch (e) {
    console.warn('notes', e);
    return slides.map(() => '');
  }
}

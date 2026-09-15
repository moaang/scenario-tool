    const WORD_XML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    function docxXmlDocument(xml) {
      const source = String(xml || '').replace(/^\s*<\?xml[^>]*\?>/, '');
      const parsed = new DOMParser().parseFromString('<scenario-word xmlns:w="' + WORD_XML_NS + '">' + source + '</scenario-word>', 'application/xml');
      if (parsed.getElementsByTagName('parsererror').length) throw new Error('Invalid DOCX XML.');
      return parsed;
    }

    function docxElements(node, name) {
      return Array.from(node.getElementsByTagNameNS(WORD_XML_NS, name));
    }

    function docxAttribute(node, name, fallback) {
      const value = node && (node.getAttributeNS(WORD_XML_NS, name) || node.getAttribute('w:' + name));
      return value == null || value === '' ? fallback : value;
    }

    function docxProperties(node, base) {
      const out = Object.assign({}, base || {});
      const first = name => docxElements(node, name)[0];
      const number = (name, attr, key, factor) => {
        const el = first(name), value = docxAttribute(el, attr, null);
        if (value !== null && Number.isFinite(Number(value))) out[key] = Number(value) * (factor || 1);
      };
      number('sz', 'val', 'fontSize', 2 / 3);
      number('spacing', 'before', 'before', 1 / 15);
      number('spacing', 'after', 'after', 1 / 15);
      number('ind', 'left', 'indent', 1 / 15);
      number('ind', 'firstLine', 'firstLine', 1 / 15);
      if (first('ind') && docxAttribute(first('ind'), 'hanging', null) !== null) out.firstLine = -Number(docxAttribute(first('ind'), 'hanging', 0)) / 15;
      const spacing = first('spacing');
      if (spacing && docxAttribute(spacing, 'line', null) !== null) {
        out.lineHeight = Number(docxAttribute(spacing, 'line', 240)) / (docxAttribute(spacing, 'lineRule', 'auto') === 'auto' ? 240 : 15);
        out.lineHeightUnit = docxAttribute(spacing, 'lineRule', 'auto') === 'auto' ? '' : 'px';
      }
      const fonts = first('rFonts');
      if (fonts) out.fontFamily = docxAttribute(fonts, 'eastAsia', docxAttribute(fonts, 'ascii', out.fontFamily || 'serif'));
      for (const [name, key] of [['b', 'bold'], ['i', 'italic'], ['pageBreakBefore', 'breakBefore'], ['keepNext', 'keepNext']]) {
        const el = first(name);
        if (el) out[key] = !['0', 'false', 'off'].includes(docxAttribute(el, 'val', '1'));
      }
      const align = first('jc');
      if (align) out.align = docxAttribute(align, 'val', 'left');
      return out;
    }

    function parseDocxPageLayout(xml) {
      const parsed = typeof xml === 'string' ? docxXmlDocument(xml) : xml;
      const section = parsed.localName === 'sectPr' ? parsed : docxElements(parsed, 'sectPr')[0];
      const layout = { pgW: 11906, pgH: 16838, marTop: 1440, marRight: 1440, marBottom: 1440, marLeft: 1440, vertical: false, linePitch: 0 };
      if (!section) return layout;
      const size = docxElements(section, 'pgSz')[0], margin = docxElements(section, 'pgMar')[0];
      layout.pgW = Number(docxAttribute(size, 'w', layout.pgW));
      layout.pgH = Number(docxAttribute(size, 'h', layout.pgH));
      for (const side of ['Top', 'Right', 'Bottom', 'Left']) layout['mar' + side] = Number(docxAttribute(margin, side.toLowerCase(), layout['mar' + side]));
      layout.orient = docxAttribute(size, 'orient', layout.pgW > layout.pgH ? 'landscape' : 'portrait');
      layout.vertical = ['tbRl', 'tbRlV', 'tbLrV'].includes(docxAttribute(docxElements(section, 'textDirection')[0], 'val', 'lrTb'));
      layout.linePitch = Number(docxAttribute(docxElements(section, 'docGrid')[0], 'linePitch', 0));
      layout.sectionBreak = docxAttribute(docxElements(section, 'type')[0], 'val', 'nextPage');
      return layout;
    }

    function parseDocxSourceModel(xml, stylesXml, commentsXml) {
      const parsed = docxXmlDocument(xml);
      const styles = docxXmlDocument(stylesXml || '');
      const defaults = docxProperties(docxElements(styles, 'docDefaults')[0] || styles, { fontSize: 16, fontFamily: 'serif' });
      const styleNodes = new Map(docxElements(styles, 'style').map(node => [docxAttribute(node, 'styleId', ''), node]));
      const styleCache = new Map();
      const resolveStyle = (id, seen) => {
        if (styleCache.has(id)) return styleCache.get(id);
        const node = styleNodes.get(id);
        if (!node || seen.has(id)) return defaults;
        seen.add(id);
        const parent = docxAttribute(docxElements(node, 'basedOn')[0], 'val', '');
        const value = docxProperties(node, parent ? resolveStyle(parent, seen) : defaults);
        styleCache.set(id, value);
        return value;
      };
      const paragraphs = [], sections = [], commentAnchors = new Map();
      let sectionStart = 0, pageBreakCount = 0;
      for (const p of docxElements(parsed, 'p')) {
        let parent = p.parentNode, omitted = false;
        while (parent) { if (['del', 'moveFrom', 'rt'].includes(parent.localName)) omitted = true; parent = parent.parentNode; }
        if (omitted) continue;
        // Nested text boxes are read at their own paragraph, not duplicated by the enclosing paragraph.
        const styleId = docxAttribute(docxElements(p, 'pStyle')[0], 'val', '');
        const direct = Array.from(p.childNodes).find(n => n.localName === 'pPr');
        const style = direct ? docxProperties(direct, resolveStyle(styleId, new Set())) : Object.assign({}, resolveStyle(styleId, new Set()));
        const paragraph = { id: 'p' + paragraphs.length, text: '', runs: [], style, section: sections.length, noteReferences: [] };
        let precedingHardBreak = false;
        const append = (text, runStyle) => {
          if (!text) return;
          paragraph.runs.push({ start: paragraph.text.length, text, style: runStyle });
          paragraph.text += text;
        };
        const walk = (node, runStyle) => {
          if (node.nodeType !== 1) return;
          const name = node.localName;
          if (['del', 'moveFrom', 'instrText', 'delInstrText', 'rt', 'pPr', 'rPr'].includes(name)) return;
          if (name === 'p' && node !== p) return;
          if (name === 't') { append(node.textContent || '', runStyle); if (node.textContent) precedingHardBreak = false; return; }
          if (name === 'footnoteReference' || name === 'endnoteReference') {
            paragraph.noteReferences.push({ kind: name.replace('Reference', ''), id: docxAttribute(node, 'id', ''), start: paragraph.text.length }); return;
          }
          if (name === 'tab') { append('\t', runStyle); return; }
          if (name === 'lastRenderedPageBreak' || name === 'br' && docxAttribute(node, 'type', '') === 'page') {
            if (name === 'lastRenderedPageBreak' && precedingHardBreak) { precedingHardBreak = false; return; }
            append('\f', runStyle); pageBreakCount += 1; precedingHardBreak = name === 'br'; return;
          }
          if (name === 'br' || name === 'cr') { append('\n', runStyle); return; }
          if (name === 'commentRangeStart' || name === 'commentReference') {
            const id = docxAttribute(node, 'id', '');
            if (!commentAnchors.has(id)) commentAnchors.set(id, { paragraphId: paragraph.id, start: paragraph.text.length });
          }
          if (name === 'commentRangeEnd') {
            const anchor = commentAnchors.get(docxAttribute(node, 'id', ''));
            if (anchor) { anchor.endParagraphId = paragraph.id; anchor.end = paragraph.text.length; }
          }
          let nextStyle = runStyle;
          if (name === 'r') {
            const properties = Array.from(node.childNodes).find(n => n.localName === 'rPr');
            if (properties) nextStyle = docxProperties(properties, runStyle);
          }
          Array.from(node.childNodes).forEach(child => walk(child, nextStyle));
        };
        walk(p, style);
        paragraphs.push(paragraph);
        const section = direct && docxElements(direct, 'sectPr')[0];
        if (section) {
          const layout = parseDocxPageLayout(section);
          sections.push(layout);
          for (let i = sectionStart; i < paragraphs.length; i += 1) paragraphs[i].section = sections.length - 1;
          sectionStart = paragraphs.length;
        }
      }
      const body = docxElements(parsed, 'body')[0];
      const finalSection = body && Array.from(body.childNodes).find(n => n.localName === 'sectPr');
      sections.push(finalSection ? parseDocxPageLayout(finalSection) : parseDocxPageLayout(parsed));
      for (let i = sectionStart; i < paragraphs.length; i += 1) paragraphs[i].section = sections.length - 1;
      const comments = docxElements(docxXmlDocument(commentsXml || ''), 'comment').map(node => {
        const id = docxAttribute(node, 'id', '');
        return { id, text: docxElements(node, 'p').map(p => docxElements(p, 't').map(t => t.textContent).join('')).join('\n'),
          author: docxAttribute(node, 'author', ''), anchor: commentAnchors.get(id) || null };
      });
      return { paragraphs, sections, comments, pageBreakCount };
    }

    async function readDocxText(file) {
      const buffer = await file.arrayBuffer();
      const [xml, styles, comments, footnotes, endnotes] = await Promise.all(['word/document.xml', 'word/styles.xml', 'word/comments.xml', 'word/footnotes.xml', 'word/endnotes.xml'].map(name => extractZipText(buffer, name)));
      if (!xml) throw new Error('word/document.xml not found in DOCX.');
      const model = parseDocxSourceModel(xml, styles, comments);
      for (const [kind, source] of [['footnote', footnotes], ['endnote', endnotes]]) {
        for (const node of docxElements(docxXmlDocument(source || ''), kind)) {
          const id = docxAttribute(node, 'id', '');
          if (Number(id) < 0 || docxAttribute(node, 'type', '') === 'separator' || docxAttribute(node, 'type', '') === 'continuationSeparator') continue;
          const paragraph = model.paragraphs.find(p => p.noteReferences.some(ref => ref.kind === kind && ref.id === id));
          if (!paragraph) continue;
          const ref = paragraph.noteReferences.find(ref => ref.kind === kind && ref.id === id);
          model.comments.push({ id: kind + '-' + id, text: docxElements(node, 'p').map(p => docxElements(p, 't').map(t => t.textContent).join('')).join('\n'),
            author: '', anchor: { paragraphId: paragraph.id, start: Math.max(0, ref.start - 1), endParagraphId: paragraph.id, end: ref.start } });
        }
      }
      const pages = await paginateDocxSource(model);
      return { text: pages.map(p => p.text).join('\n\f\n'), pageLayout: model.sections[0], model, pages };
    }

    function docxRunStyle(style) {
      const font = String(style.fontFamily || 'serif').replace(/["\\;{}<>]/g, '');
      return 'font-family:"' + font + '",serif;font-size:' + (Number(style.fontSize) || 16) + 'px;'
        + (style.bold ? 'font-weight:bold;' : '') + (style.italic ? 'font-style:italic;' : '');
    }

    function docxParagraphHtml(paragraph, start, end, layout) {
      const style = paragraph.style;
      const spacing = Number(style.lineHeight) > 0 ? Number(style.lineHeight) + (style.lineHeightUnit === 'px' ? 'px' : '') : layout.linePitch ? layout.linePitch / 15 + 'px' : '1.2';
      const runs = paragraph.runs.flatMap(run => {
        const lo = Math.max(start, run.start), hi = Math.min(end, run.start + run.text.length);
        return hi > lo ? ['<span style="' + esc(docxRunStyle(run.style)) + '">' + esc(run.text.slice(lo - run.start, hi - run.start)) + '</span>'] : [];
      }).join('');
      return '<p data-docx-paragraph="' + esc(paragraph.id) + '" data-docx-start="' + start + '" style="margin:0;white-space:pre-wrap;overflow-wrap:anywhere;line-break:strict;'
        + esc(docxRunStyle(style)) + 'line-height:' + spacing + ';text-align:' + ({ center: 'center', right: 'right', both: 'justify' }[style.align] || 'left') + ';'
        + 'text-indent:' + (start ? 0 : Number(style.firstLine) || 0) + 'px;margin-block-start:' + (start ? 0 : Number(style.before) || 0) + 'px;'
        + 'margin-block-end:' + (Number(style.after) || 0) + 'px;padding-inline-start:' + (Number(style.indent) || 0) + 'px;">' + (runs || '<br>') + '</p>';
    }

    async function paginateDocxSource(model) {
      // The browser is the layout engine. Node diagnostics expose stored break positions only.
      const canMeasure = typeof document.createRange === 'function';
      if (canMeasure && document.fonts) await document.fonts.ready;
      const pages = [];
      let page = null, host = null, content = null;
      if (canMeasure) {
        host = document.createElement('div');
        host.style.cssText = 'position:fixed;left:-100000px;top:0;visibility:hidden;contain:layout style;';
        document.body.appendChild(host);
      }
      const newPage = layout => {
        const width = layout.pgW / 15, height = layout.pgH / 15;
        page = { id: 'docx-page-' + (pages.length + 1), pageNo: pages.length + 1, sourcePageIndex: pages.length + 1,
          sourceType: 'docx', width, height, layout, text: '', textItems: [], fragments: [], annotations: [],
          paginationSource: canMeasure ? 'browser-layout' : 'stored-breaks-unmeasured' };
        pages.push(page);
        if (host) {
          host.style.width = width + 'px'; host.style.height = height + 'px';
          host.innerHTML = '<div style="position:absolute;left:' + layout.marLeft / 15 + 'px;top:' + layout.marTop / 15 + 'px;width:'
            + Math.max(1, (layout.pgW - layout.marLeft - layout.marRight) / 15) + 'px;height:' + Math.max(1, (layout.pgH - layout.marTop - layout.marBottom) / 15)
            + 'px;writing-mode:' + (layout.vertical ? 'vertical-rl' : 'horizontal-tb') + ';text-orientation:mixed;"></div>';
          content = host.firstElementChild;
        }
      };
      const fits = html => {
        if (!content) return true;
        const count = content.children.length;
        content.insertAdjacentHTML('beforeend', html);
        const b = content.getBoundingClientRect();
        const yes = Array.from(content.children).slice(count).every(element => {
          const a = element.getBoundingClientRect();
          return a.left >= b.left - 0.5 && a.right <= b.right + 0.5 && a.top >= b.top - 0.5 && a.bottom <= b.bottom + 0.5;
        });
        while (content.children.length > count) content.lastElementChild.remove();
        return yes;
      };
      const append = (paragraph, start, end, layout) => {
        const html = docxParagraphHtml(paragraph, start, end, layout);
        const text = paragraph.text.slice(start, end);
        page.fragments.push({ paragraphId: paragraph.id, start, end });
        if (page.text) page.text += '\n';
        page.text += text;
        if (!content) return;
        content.insertAdjacentHTML('beforeend', html);
        const element = content.lastElementChild;
        const bounds = host.getBoundingClientRect();
        const walker = document.createTreeWalker(element, 4);
        let node, sourceOffset = start;
        while ((node = walker.nextNode())) {
          for (let offset = 0; offset < node.textContent.length;) {
            const char = String.fromCodePoint(node.textContent.codePointAt(offset));
            const range = document.createRange(); range.setStart(node, offset); range.setEnd(node, offset + char.length);
            const rect = range.getBoundingClientRect();
            const left = (rect.left - bounds.left) / page.width * 100, top = (rect.top - bounds.top) / page.height * 100;
            const width = rect.width / page.width * 100, height = rect.height / page.height * 100;
            page.textItems.push({ str: char, text: char, pageNo: page.pageNo, leftPct: left, topPct: top, widthPct: width, heightPct: height,
              centerXPct: left + width / 2, centerYPct: top + height / 2, vertical: layout.vertical, fontSize: Number(paragraph.style.fontSize) || 16,
              source: 'docx-layout', paragraphId: paragraph.id, sourceOffset });
            sourceOffset += char.length;
            offset += char.length;
          }
        }
      };
      try {
        let lastSection = -1;
        for (let paragraphIndex = 0; paragraphIndex < model.paragraphs.length; paragraphIndex += 1) {
          const paragraph = model.paragraphs[paragraphIndex];
          const layout = model.sections[paragraph.section];
          if (!page) newPage(layout);
          else if (paragraph.section !== lastSection && lastSection >= 0) {
            const changedShape = ['pgW', 'pgH', 'marTop', 'marRight', 'marBottom', 'marLeft', 'vertical'].some(k => page.layout[k] !== layout[k]);
            if (layout.sectionBreak !== 'continuous' || changedShape) {
              if (page.fragments.length) newPage(layout);
              else { pages.pop(); newPage(layout); }
              if (layout.sectionBreak === 'evenPage' && page.pageNo % 2 || layout.sectionBreak === 'oddPage' && !(page.pageNo % 2)) newPage(layout);
            }
          }
          if (paragraph.style.breakBefore && page.fragments.length) newPage(layout);
          lastSection = paragraph.section;
          if (paragraph.style.keepNext && page.fragments.length && !paragraph.text.includes('\f')) {
            const next = model.paragraphs[paragraphIndex + 1];
            if (next && next.section === paragraph.section && !next.text.includes('\f')) {
              const pair = docxParagraphHtml(paragraph, 0, paragraph.text.length, layout) + docxParagraphHtml(next, 0, next.text.length, layout);
              if (!fits(pair)) newPage(layout);
            }
          }
          const chunks = paragraph.text.split('\f');
          let start = 0;
          for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex += 1) {
            if (chunkIndex) { newPage(layout); start += 1; }
            const limit = start + chunks[chunkIndex].length;
            if (start === limit) { if (chunks.length === 1) append(paragraph, start, limit, layout); continue; }
            while (start < limit) {
              let end = limit;
              if (!fits(docxParagraphHtml(paragraph, start, end, layout))) {
                const points = [start];
                for (const char of paragraph.text.slice(start, limit)) points.push(points[points.length - 1] + char.length);
                let lo = 0, hi = points.length - 1;
                while (lo < hi) {
                  const mid = Math.ceil((lo + hi) / 2);
                  if (fits(docxParagraphHtml(paragraph, start, points[mid], layout))) lo = mid; else hi = mid - 1;
                }
                end = points[lo];
                if (end === start) {
                  if (page.fragments.length) { newPage(layout); continue; }
                  throw new Error('DOCX page margins or font leave no space for text.');
                }
              }
              append(paragraph, start, end, layout); start = end;
              if (start < limit) newPage(layout);
            }
          }
        }
        if (!page) newPage(model.sections[0]);
        for (const comment of model.comments) {
          const anchor = comment.anchor;
          // Fragments are half-open: a comment at an automatic page boundary belongs to the next page.
          const target = pages.find(p => anchor && p.fragments.some(f => f.paragraphId === anchor.paragraphId && anchor.start >= f.start && anchor.start < f.end))
            || pages.find(p => anchor && p.fragments.some(f => f.paragraphId === anchor.paragraphId && anchor.start === f.end)) || pages[0];
          let glyphs = target.textItems.filter(item => anchor && item.paragraphId === anchor.paragraphId && item.sourceOffset >= anchor.start
            && (anchor.endParagraphId !== anchor.paragraphId || item.sourceOffset < (anchor.end == null ? anchor.start + 1 : anchor.end)));
          if (!glyphs.length && anchor) {
            const nearest = target.textItems.filter(item => item.paragraphId === anchor.paragraphId)
              .sort((a, b) => Math.abs(a.sourceOffset - anchor.start) - Math.abs(b.sourceOffset - anchor.start))[0];
            if (nearest) glyphs = [nearest];
          }
          const rect = unionOverlayRects(glyphs.map(item => ({ left: item.leftPct, top: item.topPct, width: item.widthPct, height: item.heightPct })));
          target.annotations.push({ text: comment.text, role: 'source_comment', trigger: 'docx-comment', sourceId: comment.id, author: comment.author, pageNo: target.pageNo,
            anchor, independent: true, overlayRect: rect ? Object.assign(rect, { coordinateSource: 'source-annotation', manual: false }) : null });
        }
        return pages;
      } finally { if (host) host.remove(); }
    }

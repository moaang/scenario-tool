    // DOCX source display and the PDF/DOCX shared overlay interface.
    // Source paragraphs draw the page background; prepared block geometry drives the shared
    // overlay and revision-note composer. Older saved documents can use their existing blocks.
    function docxReaderBlockKind(block) {
      if (block.type === 'end_marker' || block.zone === 'body_end') return 'end';
      if (block.type === 'scene_heading') return 'scene';
      if (block.type === 'dialog' || block.speaker) return 'dialog';
      if (block.zone === 'cast_list' || block.zone === 'front_matter'
        || FRONT_AND_CAST_BLOCK_TYPES.includes(block.type)) return 'cast';
      return 'action';
    }

    function renderDocxReaderBlock(block) {
      const kind = docxReaderBlockKind(block);
      let inner;
      if (kind === 'dialog') {
        // rawText already carries the 「」 quotes and any （modifier）; just tint the speaker name
        // rather than re-wrapping (which double-quoted the line and dropped the modifier).
        const speaker = block.speaker || block.characterName || '';
        const raw = displayPdfBlockText(String(block.rawText || block.text || ''));
        const style = characterColorStyle(speaker || 'unknown-speaker');
        inner = (speaker && raw.startsWith(speaker))
          ? `<span class="dvr-speaker" style="${style}">${esc(speaker)}</span>${esc(raw.slice(speaker.length))}`
          : esc(raw);
      } else {
        inner = esc(String(block.text || block.rawText || '')).replace(/\n+/g, '　');
      }
      // Each block sits on its own vertical line/column (how a printed vertical script reads:
      // one speech / one direction per line, right to left); scene headings get an extra blank
      // line before them as an act break. Selection/notes are drawn by the overlay boxes + the
      // PDF feedback layer on top, so the text itself is just the page background here.
      const lead = kind === 'scene' ? '<br>' : '';
      return `${lead}<span class="dvr-block dvr-${kind}" data-dvr-block="${esc(block.id)}">${inner}</span><br>`;
    }

    function docxReaderRenderKey(doc, page, pages, blocks, body, metrics) {
      return JSON.stringify([doc.id, page.pageNo, pages, page.layout || doc.docxPageLayout, body, metrics,
        blocks, blocks.map(b => characterColorStyle(b.speaker || b.characterName || 'unknown-speaker')),
        t('docxReaderHint'), t('docxAttachPdf')]);
    }

    function docxPageDisplayMetrics(layout, availWidth) {
      const pageW = Math.max(240, Math.min(Number(availWidth) || 900, 1080));
      const scale = pageW / (layout.pgW / 15);
      return { pageW, pageH: layout.pgH / 15 * scale, scale };
    }

    function docxPageSourceHtml(doc, page, pageBlocks) {
      const paragraphs = new Map((doc.docxSource && doc.docxSource.paragraphs || []).map(p => [p.id, p]));
      if (paragraphs.size && page.fragments) return page.fragments.map(fragment => {
        const paragraph = paragraphs.get(fragment.paragraphId);
        return paragraph ? docxParagraphHtml(paragraph, fragment.start, fragment.end, page.layout) : '';
      }).join('');
      // Older saved documents have blocks and page spans but no original OOXML model.
      return pageBlocks.map(block => renderDocxReaderBlock(Object.assign({}, block,
        { text: overlayTextForBlockPage(block, page.pageNo) || block.text }))).join('');
    }

    function renderDocxVerticalReader(el, doc) {
      const defaultLayout = doc.docxPageLayout || parseDocxPageLayout('');
      const pages = doc.docxPages && doc.docxPages.length ? doc.docxPages : sourceReviewPagesFromText(doc.originalText || '')
        .map(p => ({ pageNo: p.pageNo, layout: defaultLayout, text: p.lines.join('\n') }));
      const pageNos = pages.map(page => page.pageNo);
      const activePageNo = docxReaderActivePageNo(pageNos);
      const page = pages.find(p => p.pageNo === activePageNo) || pages[0];
      const layout = page.layout || defaultLayout;
      const index = pageNos.indexOf(activePageNo);
      const pageBlocks = sortedDocumentBlocks(doc).filter(block => blockHasOverlayPage(block, activePageNo) && isOverlayEligibleBlock(block));
      const metrics = docxPageDisplayMetrics(layout, (el.clientWidth || 900) - 64);
      const body = docxPageSourceHtml(doc, page, pageBlocks);
      const renderKey = docxReaderRenderKey(doc, page, pageNos, pageBlocks, body, metrics);
      if (el.dataset.renderKey === renderKey && el.querySelector('.docx-page')) {
        syncBodyOverlaySelectionClasses();
        syncPdfOverlaySelectionToolbar();
        renderPdfFeedbackLayer();
        return;
      }
      closePdfFeedbackComposer();
      const contentStyle = 'position:absolute;left:' + layout.marLeft / 15 + 'px;top:' + layout.marTop / 15 + 'px;width:'
        + Math.max(1, (layout.pgW - layout.marLeft - layout.marRight) / 15) + 'px;height:' + Math.max(1, (layout.pgH - layout.marTop - layout.marBottom) / 15)
        + 'px;writing-mode:' + (layout.vertical ? 'vertical-rl' : 'horizontal-tb') + ';text-orientation:mixed;';
      el.innerHTML = `<div class="docx-reader-wrap"><div class="docx-reader-bar" id="docxReaderBar">
        <button class="small" data-docx-page="prev" ${index <= 0 ? 'disabled' : ''}>‹</button>
        <span class="drb-kind">Page ${activePageNo} / ${pages.length}</span>
        <button class="small" data-docx-page="next" ${index >= pages.length - 1 ? 'disabled' : ''}>›</button>
        <span class="drb-spacer"></span><span class="drb-hint">${esc(t('docxReaderHint'))}</span>
        <button class="small" data-docx-pdf>${esc(t('docxAttachPdf'))}</button><input type="file" accept=".pdf" data-docx-pdf-input hidden></div>
        <div class="docx-pages-scroll"><div class="docx-pages">
        <div class="docx-page pdf-page-surface" id="pdfPageSurface" data-page-no="${activePageNo}" style="width:${metrics.pageW}px;height:${metrics.pageH}px;min-height:0;background:white;">
          <div style="position:absolute;left:0;top:0;width:${layout.pgW / 15}px;height:${layout.pgH / 15}px;transform-origin:0 0;transform:scale(${metrics.scale});">
            <div class="docx-source-content" style="${contentStyle}">${body}</div>
          </div>
          <div class="pdf-overlay-layer" id="pdfOverlayLayer" data-pdf-overlay-layer></div>
          <div class="pdf-feedback-layer" id="pdfFeedbackLayer" data-pdf-feedback-layer></div>
        </div></div></div></div>`;
      el.dataset.renderKey = renderKey;
      const bar = el.querySelector('#docxReaderBar');
      const input = bar.querySelector('[data-docx-pdf-input]');
      bar.querySelector('[data-docx-pdf]').addEventListener('click', () => input.click());
      input.addEventListener('change', async () => {
        const file = input.files && input.files[0];
        if (!file) return;
        try { await attachDocxLayoutPdf(doc, file); } catch (error) { alert(error.message || String(error)); }
      });
      bar.addEventListener('pointerdown', event => {
        const button = event.target.closest('[data-docx-page]');
        if (!button) return;
        const next = index + (button.dataset.docxPage === 'prev' ? -1 : 1);
        if (pageNos[next] == null) return;
        state.activeScriptPageNo = pageNos[next];
        renderScriptStructurePanel();
      });
      buildDocxOverlayBoxes(el, pageBlocks, activePageNo);
      renderPdfFeedbackLayer();
    }

    async function attachDocxLayoutPdf(doc, file) {
      const result = await readDocumentFileSelfContained(file);
      if (result.type !== 'pdf' || result.warning) throw new Error(result.warning || t('docxPdfMismatch'));
      const candidate = Object.assign({}, doc, { draftSourceType: 'pdf', draftPages: result.pages,
        pdfAnnotations: result.pdfAnnotations, originalText: result.text, rawImportedText: result.text });
      candidate.reviewItems = buildSourceReviewItemsForDocument(result.text, candidate);
      const mapped = normalizeImportedBlocks(sourceReviewItemsToBlocks(candidate.reviewItems, candidate));
      const contentBlocks = list => list.filter(b => b.type !== 'pdf_note' && b.type !== 'pdf_link');
      const old = sortedDocumentBlocks(doc), oldContent = contentBlocks(old), nextContent = contentBlocks(mapped);
      const key = b => b.type + '|' + normalizePdfCoordinateMatchText(b.rawText || b.text || '');
      if (oldContent.length !== nextContent.length || oldContent.some((b, i) => key(b) !== key(nextContent[i]))) throw new Error(t('docxPdfMismatch'));
      // Validate the entire mapping before modifying document state or any user annotation.
      if (activeDocument().id !== doc.id) throw new Error(t('docxPdfMismatch'));
      const oldNotes = old.filter(b => b.sourceAnnotation && !b.sourceAnnotation.href);
      const noteMappings = oldNotes.map(note => {
        const rect = overlayRectForBlockPage(note, note.pageNo);
        const matches = oldContent.map((block, index) => ({ index, area: blockHasOverlayPage(block, note.pageNo)
          ? overlayRectOverlapArea(rect, overlayRectForBlockPage(block, note.pageNo)) : 0 })).filter(item => item.area > 0);
        matches.sort((a, b) => b.area - a.area);
        return { note, next: matches.length ? nextContent[matches[0].index] : null };
      });
      saveActiveDocumentVersionSnapshot(doc);
      for (const { note, next } of noteMappings) {
        if (!next || !next.overlayRect) continue;
        note.previousLayoutGeometry = { pageNo: note.pageNo, overlayRect: cloneJson(note.overlayRect), pageSpans: cloneJson(note.pageSpans || []) };
        note.pageNo = next.pageNo;
        note.overlayRect = Object.assign({}, next.overlayRect, { coordinateSource: 'source-annotation' });
        note.pageSpans = [{ pageNo: note.pageNo, text: note.rawText, overlayRect: cloneJson(note.overlayRect) }];
        note.sourceAnnotation.pageNo = note.pageNo; note.sourceAnnotation.overlayRect = cloneJson(note.overlayRect);
      }
      const additions = mapped.filter(b => (b.type === 'pdf_note' || b.type === 'pdf_link')
        && !old.some(previous => previous.type === b.type && normalizePdfCoordinateMatchText(previous.rawText) === normalizePdfCoordinateMatchText(b.rawText)));
      additions.forEach((block, i) => {
        block.id = doc.id + '-layout-note-' + hashText(file.name + '|' + i + '|' + block.rawText);
        block.sourceDocumentId = doc.id;
      });
      state.project.sourceBlocks.push(...additions);
      syncSourceAnnotationNotesForDocument(state.project, doc);
      oldContent.forEach((block, i) => {
        const next = nextContent[i];
        const previous = { pageNo: block.pageNo, overlayRect: cloneJson(block.overlayRect), pageSpans: cloneJson(block.pageSpans || []) };
        block.previousLayoutGeometry = previous;
        block.pageNo = next.pageNo; block.pageSpans = cloneJson(next.pageSpans || []); block.overlayRect = cloneJson(next.overlayRect);
        copySourceReviewGeometry(block, previous);
      });
      doc.sourceFormat = 'docx'; doc.layoutPdfFileName = file.name; doc.draftSourceType = 'pdf';
      doc.draftPages = result.pages; doc.pageMetadata = sourcePagesToPageMetadata(result.pages);
      doc.pdfAnnotations = (doc.pdfAnnotations || []).concat(result.pdfAnnotations || []);
      doc.pdfExtractionSignals = result.pdfExtractionSignals;
      clearDocumentPdfPreview(doc); doc.pdfCacheKey = ''; doc.pdfCacheMeta = null;
      setPdfPreviewFileForDocument(doc, file);
      await putPdfPreviewFileCache(doc, file).catch(error => console.warn('PDF cache save failed', error));
      createNewDocumentVersionSnapshot(doc, { label: t('docxAttachPdf'), force: true });
      markProjectChanged('attach_docx_layout_pdf');
      state.activeScriptPageNo = 1; renderScriptStructurePanel();
    }

    function buildDocxOverlayBoxes(el, pageBlocks, pageNo) {
      const surface = el.querySelector('#pdfPageSurface'), overlay = el.querySelector('#pdfOverlayLayer');
      if (!surface || !overlay) return;
      // Only old saved documents lack import-time geometry. Their original block DOM is measured once.
      const bounds = surface.getBoundingClientRect();
      if (bounds.width && bounds.height) el.querySelectorAll('.dvr-block').forEach(span => {
        const block = pageBlocks.find(b => b.id === span.getAttribute('data-dvr-block'));
        if (!block || overlayRectForBlockPage(block, pageNo)) return;
        const rect = span.getBoundingClientRect();
        setBlockOverlayRectForPage(block, pageNo, { left: (rect.left - bounds.left) / bounds.width * 100,
          top: (rect.top - bounds.top) / bounds.height * 100, width: rect.width / bounds.width * 100,
          height: rect.height / bounds.height * 100, coordinateSource: 'docx-layout', manual: false });
      });
      const doc = activeDocument();
      const proxies = pageBlocks.map(block => pdfOverlayBlockForPage(block, pageNo)).filter(Boolean);
      overlay.innerHTML = proxies.map(block => renderBodyOverlayBox(doc, block)).join('');
      bindBodyOverlayInteractions();
    }

    function docxReaderActivePageNo(pageNos) {
      const want = Number(state.activeScriptPageNo);
      return pageNos.includes(want) ? want : (pageNos[0] || 1);
    }

    const pdfDisplayLoads = new WeakMap();
    const pdfCanvasRenderTasks = new WeakMap();

    async function renderPdfJsPageForActiveBodyPreview(doc, pageNo) {
      if (state.activeMiddleTab !== 'script') return;
      const file = doc && state.pdfPreviewFiles && state.pdfPreviewFiles[doc.id];
      const canvas = $('pdfPageCanvas');
      const surface = $('pdfPageSurface');
      const status = $('pdfJsStatus');
      if (!file || !canvas || !surface) return;
      const token = ++state.pdfRenderToken;
      clearPdfTextFallbackPreview(surface);
      if (status) status.textContent = 'PDF.js...';
      const pdfjsLib = await loadPdfJsLibrary();
      let pdf = state.pdfJsDocuments && state.pdfJsDocuments[doc.id];
      if (!pdf) {
        if (!pdfDisplayLoads.has(file)) pdfDisplayLoads.set(file, (async () => {
          const data = await file.arrayBuffer();
          return await pdfjsLib.getDocument(pdfJsDocumentParams(data)).promise;
        })());
        try { pdf = await pdfDisplayLoads.get(file); }
        catch (error) { pdfDisplayLoads.delete(file); throw error; }
        if (state.pdfPreviewFiles[doc.id] !== file) { await pdf.destroy(); return; }
        state.pdfJsDocuments[doc.id] = pdf;
      }
      if (token !== state.pdfRenderToken) return;
      const page = await pdf.getPage(Math.max(1, Math.min(Number(pageNo) || 1, pdf.numPages || 1)));
      const stage = surface.closest('.pdf-preview-stage');
      const baseViewport = page.getViewport({ scale: 1 });
      // 여백 48 = .pdf-page-scroll 의 좌우 패딩 24px 두 벌(가운데 패널은 그 패딩을 0 으로 덮는다).
      const availableWidth = Math.max(420, (stage && stage.clientWidth || 920) - 48);
      // Fit the preview to its pane within the supported display scale bounds.
      const scale = Math.max(0.65, Math.min(2.6, availableWidth / baseViewport.width));
      const viewport = page.getViewport({ scale });
      // Render the raster at a density that follows the canvas zoom so magnifying the
      // shell (a CSS transform) doesn't upscale a fixed bitmap into mush. Capped + stepped
      // to avoid re-render churn; the CSS display size stays viewport.width.
      const dpr = pdfRenderDensity();
      const ctx = canvas.getContext('2d');
      const renderKey = [doc.id, pageNo, Math.round(viewport.width), Math.round(viewport.height), Math.round(dpr * 100)].join('#');
      const targetWidth = Math.floor(viewport.width * dpr);
      const targetHeight = Math.floor(viewport.height * dpr);
      canvas.style.width = Math.floor(viewport.width) + 'px';
      canvas.style.height = Math.floor(viewport.height) + 'px';
      surface.style.width = Math.floor(viewport.width) + 'px';
      surface.style.height = Math.floor(viewport.height) + 'px';
      surface.style.minHeight = Math.floor(viewport.height) + 'px';
      let pending = pdfCanvasRenderTasks.get(canvas);
      if (pending && pending.key !== renderKey) {
        pending.task.cancel();
        try { await pending.task.promise; } catch (error) { if (error.name !== 'RenderingCancelledException') throw error; }
        pending = null;
      }
      if (token !== state.pdfRenderToken) return;
      if (canvas.dataset.renderKey !== renderKey || canvas.width !== targetWidth || canvas.height !== targetHeight) {
        if (!pending) {
          canvas.width = targetWidth; canvas.height = targetHeight;
          pending = { key: renderKey, task: page.render({ canvasContext: ctx, viewport, transform: dpr === 1 ? null : [dpr, 0, 0, dpr, 0, 0] }) };
          pdfCanvasRenderTasks.set(canvas, pending);
        }
        try { await pending.task.promise; }
        catch (error) { if (error.name === 'RenderingCancelledException') return; throw error; }
        finally { if (pdfCanvasRenderTasks.get(canvas) === pending) pdfCanvasRenderTasks.delete(canvas); }
        if (token !== state.pdfRenderToken) return;
        canvas.dataset.renderKey = renderKey;
      }
      const items = await preparePdfPageForDisplay(doc, pageNo, page, pdfjsLib);
      if (token !== state.pdfRenderToken) return;
      const usingFallbackTextLayer = pdfTextItemsCoordinateSource(items) === 'pdf-logical-fallback';
      if (usingFallbackTextLayer) renderPdfFallbackTextLayer(items, pageNo);
      else clearPdfFallbackTextLayer();
      renderPdfLayoutGuideLayer(items, pageNo);

      if (status) {
        status.textContent = (usingFallbackTextLayer ? t('pdfPreviewFallback') : t('pdfPreviewNative')) + ' / textLayer ' + items.length;
      }
    }

    async function preparePdfPageForDisplay(doc, pageNo, page, pdfjsLib) {
      const key = doc.id + '#' + pageNo;
      if (state.pdfPageTextItems[key]) return state.pdfPageTextItems[key];
      const sourcePage = pdfLogicalPageForDocument(doc, pageNo);
      let items = sourcePage && sourcePage.textItems;
      if (!items || !items.length) {
        const content = await page.getTextContent({ includeMarkedContent: true });
        items = await pdfTextItemsForPage(page, content, page.getViewport({ scale: 1 }), pdfjsLib, pageNo);
        if (shouldUsePdfLogicalFallbackTextItems(items, sourcePage)) items = pdfLogicalFallbackTextItems(sourcePage, pageNo);
        annotatePdfJsTextItemsFromLogicalAnnotations(items, pdfAnnotationsForRenderedPdfPage(doc, pageNo));
      }
      state.pdfPageTextItems[key] = items;
      // Stored older projects can lack extracted pages. Repair their missing geometry once.
      ensurePdfOverlayRectsForPageFromDraft(doc, pageNo);
      ensurePdfOverlayRectsForRenderedPage(doc, pageNo, items);
      return items;
    }

    async function readDocumentFileSelfContained(file) {
      beginPdfOverlayDropImport();
      let result;
      try {
        result = await readDocumentFileSelfContainedInner(file);
        return result;
      } finally {
        const importDrops = endPdfOverlayDropImport();
        if (result) result.pdfExtractionSignals = Object.assign(
          pdfReviewSignalsFromAnnotations(result.pdfAnnotations || pdfAnnotationsFromPages(result.pages || [])),
          result.pdfExtractionSignals || {}, { importDrops });
      }
    }

    async function readDocumentFileSelfContainedInner(file) {
      const name = String(file && file.name || '').toLowerCase();
      if (name.endsWith('.md') || name.endsWith('.markdown')) {
        const markdownSource = await file.text();
        const markdown = parseMarkdownDocument(markdownSource);
        return {
          type: 'md',
          text: markdown.text,
          markdownSource,
          markdownBlocks: markdown.markdownBlocks,
          blocks: markdown.parsed.blocks,
          warning: ''
        };
      }
      if (name.endsWith('.txt') || name.endsWith('.text')) {
        const text = await file.text();
        return { type: 'txt', text, warning: '' };
      }
      if (name.endsWith('.docx')) {
        const docxResult = await readDocxText(file);
        return { type: 'docx', text: docxResult.text, docxPageLayout: docxResult.pageLayout,
          pages: docxResult.pages, docxSource: docxResult.model,
          docxPages: docxResult.pages.map(page => Object.assign({}, page, { textItems: undefined })), warning: '' };
      }
      if (name.endsWith('.pdf')) {
        return readPdfDocumentWithPdfJs(file);
      }
      const text = await file.text();
      return { type: 'txt', text, warning: 'Unknown extension read as text.' };
    }

    // Tagged content can supply logical order; PDF.js glyph coordinates supply physical lines.
    // readPdfDocumentWithPdfJs chooses tagged text or a selective stream-parser fallback per page.
    function pdfStructuredLogicalPage(tree, content, pageNo) {
      if (!tree) return null;
      const byId = new Map(), stack = [], artifacts = [];
      let bodyCharacters = 0;
      for (const item of content.items || []) {
        if (item.type === 'beginMarkedContent' || item.type === 'beginMarkedContentProps') {
          stack.push({ id: item.id, tag: item.tag });
          if (item.id && !byId.has(item.id)) byId.set(item.id, '');
          continue;
        }
        if (item.type === 'endMarkedContent') { stack.pop(); continue; }
        if (typeof item.str !== 'string' || !item.str.trim()) continue;
        if (stack.some(mark => mark.tag === 'Artifact')) {
          artifacts.push({ text: item.str, x: item.transform[4], y: item.transform[5] }); continue;
        }
        bodyCharacters += normalizePdfCompactTextNfkc(item.str).length;
        const mark = stack.slice().reverse().find(entry => entry.id);
        if (mark) byId.set(mark.id, (byId.get(mark.id) || '') + item.str);
      }
      const paragraphs = [], used = new Set();
      let missing = false, duplicate = false;
      const collect = (node, ids) => {
        if (node.type === 'content') {
          if (!byId.has(node.id)) { missing = true; return ''; }
          if (used.has(node.id)) { duplicate = true; return ''; }
          used.add(node.id); ids.push(node.id); return byId.get(node.id);
        }
        return (node.children || []).map(child => collect(child, ids)).join('');
      };
      const visit = node => {
        if (['P', 'H', 'H1', 'H2', 'H3', 'LI', 'TD', 'TH'].includes(node.role)) {
          const ids = [], text = collect(node, ids);
          if (text.trim()) paragraphs.push({ role: node.role, text, markIds: ids });
        } else (node.children || []).forEach(visit);
      };
      visit(tree);
      const text = paragraphs.map(p => p.text).join('\n');
      const covered = normalizePdfCompactTextNfkc(text).length;
      const coverage = bodyCharacters ? covered / bodyCharacters : 1;
      if (missing || duplicate || !paragraphs.length || coverage < 0.995 || coverage > 1.005 || /[\ufffd\u0000]/.test(text)) return null;
      const printed = pdfPrintedPageNoItems(artifacts).map(item => item.text).join('');
      return { pageNo, text, printedPageNo: printed ? Number(printed.normalize('NFKC')) : null, annotations: [],
        requiresLayoutNotes: paragraphs.some(p => isPdfDetachedNoteText(p.text) || p.text.includes('※')) || artifacts.some(item => isPdfDetachedNoteText(item.text)),
        structure: { coverage, paragraphs }, logicalSource: 'pdf-structure' };
    }

    function pdfStandardAnnotations(annotations, viewport, pageNo) {
      return (annotations || []).flatMap(note => {
        const text = String(note.contentsObj && note.contentsObj.str || note.contents || '').trim();
        const rawHref = String(note.url || note.unsafeUrl || '').trim();
        const href = /^(https?:|mailto:)/i.test(rawHref) ? rawHref : '';
        if (!text && !href) return [];
        if (note.subtype === 'Popup' || note.fieldType) return [];
        let overlayRect = null;
        if (Array.isArray(note.rect) && note.rect.length === 4) {
          const r = viewport.convertToViewportRectangle(note.rect);
          overlayRect = { left: Math.min(r[0], r[2]) / viewport.width * 100, top: Math.min(r[1], r[3]) / viewport.height * 100,
            width: Math.abs(r[2] - r[0]) / viewport.width * 100, height: Math.abs(r[3] - r[1]) / viewport.height * 100,
            manual: false, inferred: false, coordinateSource: 'source-annotation' };
        }
        return [{ id: 'source-' + pageNo + '-' + note.id, sourceId: note.id, text: text || href, href,
          pageNo, role: href ? 'pdf_link' : 'source_comment', trigger: href ? 'link' : 'pdf-annotation', overlayRect,
          author: note.titleObj && note.titleObj.str || '', independent: true }];
      });
    }

    async function readPdfDocumentWithPdfJs(file) {
      const data = await file.arrayBuffer();
      let logicalPages = null;
      let logicalPageByNo = new Map();
      const readLogicalFallback = async selectedPages => {
        if (logicalPages === null) logicalPages = await readPdfLogicalPagesFromBuffer(data, selectedPages);
        (logicalPages || []).forEach(page => logicalPageByNo.set(Number(page.pageNo || page.sourcePageIndex || 0), page));
      };
      let pdfjsLib;
      let pdf;
      try {
        pdfjsLib = await loadPdfJsLibrary();
        pdf = await pdfjsLib.getDocument(pdfJsDocumentParams(data.slice(0))).promise;
      } catch (error) {
        await readLogicalFallback();
        if (logicalPages && logicalPages.length) {
          return pdfDocumentResultFromLogicalPages(logicalPages, error);
        }
        throw error;
      }
      try {
      const pages = [];
      const totalPageCount = Number(pdf.numPages) || 0;
      const maxPages = Math.min(totalPageCount, PDF_MAX_PAGES_TO_PARSE);
      // 상한에 걸리면 뒷부분이 통째로 빠진 채 정상 결과처럼 반환된다 — 조용히 넘기지 않는다(자체 파서의 같은 상한도 pageBudgetExceeded 로 남긴다).
      const cappedPageCount = Math.max(0, totalPageCount - maxPages);
      if (cappedPageCount) notePdfOverlayNote('pdfJsPageCapExceeded', 'p' + (maxPages + 1) + '+ x' + cappedPageCount);
      // Extract text, structure, annotations and layout signals once for each page.
      // 상한은 문서 전체의 괘선을 봐야 정해진다. 예전에는 줄 묶기가 그 쪽 후보값을 쓰고
      // 확정은 파싱이 다 끝난 뒤에 나서, **같은 판정이 두 값으로 갈라져 있었다.**
      // 여기서 한 번 정해 아래 모든 층이 같은 값을 쓴다. getPage 는 pdf.js 가 캐시하므로
      // 두 패스로 나눠도 문서를 두 번 읽지 않는다.
      const pageOperatorSignals = [];
      const pageInputs = [];
      for (let pageNo = 1; pageNo <= maxPages; pageNo += 1) {
        const page = await pdf.getPage(pageNo);
        const viewport = page.getViewport({ scale: 1 });
        const [signals, textContent, tree, annotations] = await Promise.all([
          pdfJsPageOperatorSignals(page, pdfjsLib, viewport).catch(() => ({ horizontalDividers: [] })),
          page.getTextContent({ includeMarkedContent: true }), page.getStructTree().catch(() => null),
          page.getAnnotations().catch(() => [])
        ]);
        pageOperatorSignals.push(signals);
        pageInputs.push({ textContent, logicalPage: pdfStructuredLogicalPage(tree, textContent, pageNo),
          annotations: pdfStandardAnnotations(annotations, viewport, pageNo) });
        reportPdfImportProgress(pageNo, maxPages * 2);
        if (pageNo % 3 === 0) await yieldToBrowser();
      }
      const fallbackPageNos = new Set();
      pageInputs.forEach((input, index) => {
        if (!input.logicalPage || input.logicalPage.requiresLayoutNotes || pageOperatorSignals[index].hasColorText) fallbackPageNos.add(index + 1);
        else logicalPageByNo.set(index + 1, input.logicalPage);
      });
      if (fallbackPageNos.size) await readLogicalFallback(fallbackPageNos);
      const stableBodyTopLinePct = pdfJsStableBodyTopLineFromCandidates(
        pageOperatorSignals.map(signals =>
          pdfJsBodyTopLineFromLayoutDividers(signals && signals.horizontalDividers || [])),
        maxPages);

      // Derive page lines and body ranges from the extracted data.
      for (let pageNo = 1; pageNo <= maxPages; pageNo += 1) {
        const page = await pdf.getPage(pageNo);
        const viewport = page.getViewport({ scale: 1 });
        const textContent = pageInputs[pageNo - 1].textContent;
        const operatorSignals = pageOperatorSignals[pageNo - 1];
        const layoutDividers = operatorSignals && operatorSignals.horizontalDividers || [];
        const logicalPage = logicalPageByNo.get(pageNo) || null;
        let textItems = await pdfTextItemsForPage(page, textContent, viewport, pdfjsLib, pageNo);
        const usingLogicalFallbackItems = shouldUsePdfLogicalFallbackTextItems(textItems, logicalPage);
        if (usingLogicalFallbackItems) {
          textItems = pdfLogicalFallbackTextItems(logicalPage, pageNo);
        }
        annotatePdfJsTextItemsFromLogicalAnnotations(textItems, logicalPage && logicalPage.annotations);
        const preSplitLogicalText = normalizePdfLogicalPageText(logicalPage && logicalPage.text);
        // 본문 상한(괘선) 위에는 가로쓰기 지문이 놓인다. 그것을 세로 묶기에 함께 넣으면
        // 가로줄 여러 개를 세로로 가로지르며 글자를 주워 담아 뜻 없는 줄이 나온다 — 실측:
        // 그런 줄이 한 문서에서만 71 건이었고, 꼬리 복구가 「근거 0건」으로 흘리던 것의
        // 대부분이었다. 다만 **상한 위에는 씬 번호도 있다**(실측: 지문 w7~22 · 씬번호 w1.31).
        // 그래서 위치로 범위를 좁힌 뒤 **모양**으로 가른다 — 위치만 보면 씬 번호가 본문에서
        // 사라지고, 모양만으로 전 페이지를 훑으면 본문 안의 넓은 아이템까지 걸려 줄을 끊는
        // 간격 통계가 흔들린다(둘 다 실측으로 확인했다).
        const bodyTopCandidatePct = pdfJsPageBodyTopLinePct(
          pdfJsBodyTopLineFromLayoutDividers(layoutDividers), stableBodyTopLinePct);
        const aboveTopHorizontalItems = bodyTopCandidatePct == null ? []
          : textItems.filter(item => item
            && (Number(item.centerYPct) || 0) < bodyTopCandidatePct
            && isClearlyHorizontalRunItem(item));
        const aboveTopItemSet = new Set(aboveTopHorizontalItems);
        const columnItems = aboveTopHorizontalItems.length
          ? textItems.filter(item => !aboveTopItemSet.has(item))
          : textItems;
        const allLines = groupPdfJsTextItemsToOverlayLines(columnItems, pageNo)
          .concat(aboveTopHorizontalItems.length
            ? groupPdfJsTextItemsHorizontal(aboveTopHorizontalItems) : [])
          .filter(line => keepPdfOverlayLineNotingArtifactDrop(line, pageNo));
        const bodyCandidateContext = createPdfOverlayBodyCandidateContext(allLines, pageNo, preSplitLogicalText);
        const lineGroups = splitPdfJsOverlayLinesForBodyAndAnnotations(allLines, pageNo, { bodyCandidateContext });
        const bodyCutLineCandidatePct = pdfJsBodyCutLineFromColorLines(lineGroups.colorLines)
          || pdfJsBodyCutLineFromDetachedLowerLines(lineGroups.bodyLines, pageNo);
        const bodyTopLineCandidatePct = pdfJsBodyTopLineFromLayoutDividers(layoutDividers);
        const mergedAnnotations = mergePdfPageAnnotations(lineGroups.annotations, logicalPage && logicalPage.annotations)
          .concat(pageInputs[pageNo - 1].annotations);
        const rawLogicalText = removePdfColorAnnotationLinesFromLogicalText(preSplitLogicalText, mergedAnnotations);
        const rubyLines = pdfRubyLinesFromTextItems(textItems);
        const bodyLines = mergeDetachedPdfBodyLines(expandPdfOverlayLineBoundsForMissingLogicalTails(lineGroups.bodyLines, rawLogicalText, { pageNo }));
        lineGroups.bodyLines = bodyLines;
        const horizontalMetadataLines = pdfJsHorizontalMetadataLinesForBodyItems(textItems, bodyLines, pageNo);
        // 행번호는 아이템 단계에서 판정한다. 예전에는 여기서 pdf.js 오버레이 라인을 오라클로
        // 삼아 자체 파서 텍스트의 선두 숫자를 다시 깎았는데, 그것이 두 검출기의 불일치를 텍스트로
        // 되먹이는 통로였다 — 선두가 '0' 이면 페이지의 아무 줄과 8자 접두사만 맞아도 지웠고,
        // 「오버레이 인덱스 == 인쇄 행번호 − 1」이라는 가정은 위의 아티팩트 필터가 이미 깨뜨린다.
        // 중간 변수는 남기지 않는다 — 소비자가 하나뿐이었다(같은 값에 두 이름을 두지 않는다).
        const logicalText = removePdfRubyLinesFromLogicalText(rawLogicalText, rubyLines);
        // 읽는 사람이 보는 페이지 텍스트는 여기서 두 번 더 깎인다(파란 주석 줄, 루비 줄).
        // **두 단계 다 몇 줄을 지웠는지 기록하지 않는다** — 줄이 통째로 없어져도 화면에는
        // 그저 글자가 없을 뿐이라 어디서 잃었는지 알 수가 없다. 기록을 안 붙인 근거는 불명이다.
        pages.push({
          id: 'pdfjs-page-' + pageNo,
          pageNo,
          printedPageNo: pdfPrintedPageNoValue(logicalPage && logicalPage.printedPageNo),
          sourcePageIndex: pageNo,
          width: viewport.width,
          height: viewport.height,
          textItems,
          lines: bodyLines,
          bodyBaseLines: bodyLines,
          colorLines: lineGroups.colorLines,
          horizontalMetadataLines,
          pdfJsMergedBodyLines: lineGroups.bodyLines,
          bodyCutLineCandidatePct,
          bodyTopLineCandidatePct,
          layoutDividers,
          rawLogicalText,
          logicalText,
          structure: pageInputs[pageNo - 1].logicalPage && pageInputs[pageNo - 1].logicalPage.structure || null,
          // bodyReviewText / text / bodyCutLinePct / bodyTopLinePct 는 여기서 쓰지 않는다 —
          // applyStablePdfJsBodyCutLineToPages 가 이 객체가 밖으로 나가기 전에 전부 덮어쓴다.
          // 두 벌을 두면 반드시 갈라진다(logicalLines 가 그렇게 아무도 못 읽는 사본이 됐다).
          textSource: usingLogicalFallbackItems ? 'pdf-logical-fallback' : (logicalText ? 'pdf-logical' : 'pdfjs-overlay'),
          annotations: mergedAnnotations
        });
        reportPdfImportProgress(maxPages + pageNo, maxPages * 2);
        if (pageNo % 3 === 0) await yieldToBrowser();
      }
      applyStablePdfJsBodyCutLineToPages(pages, stableBodyTopLinePct);
      const pdfAnnotations = pdfAnnotationsFromPages(pages);
      const pdfSignals = pdfReviewSignalsFromAnnotations(pdfAnnotations);
      pdfSignals.logicalFallbackUsed = logicalPages !== null;
      pdfSignals.logicalFallbackPageCount = fallbackPageNos.size;
      pdfSignals.structuredPageUsedCount = maxPages - fallbackPageNos.size;
      pdfSignals.structuredPageCount = pageInputs.filter(input => input.logicalPage).length;
      pdfSignals.layoutDividerCount = (pages || []).reduce((sum, page) => sum + ((page && page.layoutDividers || []).length), 0);
      pdfSignals.bodyTopLinePageCount = (pages || []).filter(page => page && isPdfBodyTopLineTrigger(page.bodyTopLineTrigger)).length;
      return {
        type: 'pdf',
        text: pages.map(page => page.text || '').join('\n\f\n'),
        pages,
        pdfAnnotations,
        pdfExtractionSignals: pdfSignals,
        warning: cappedPageCount
          ? 'PDF page cap reached: parsed ' + maxPages + ' of ' + totalPageCount + ' pages; ' + cappedPageCount + ' pages were not read.'
          : ''
      };
      } finally {
        // pdf.js keeps per-document caches (fonts, decoded images, page structures) alive in its
        // worker thread until destroy() is called — dropping the JS reference alone does not free
        // them. Without this, every PDF import (drag-drop or re-import) leaks worker memory for
        // the life of the tab; destroy() here covers both the success and error-during-parse paths.
        try { pdf.destroy(); } catch (error) {}
      }
    }

    function pdfDocumentResultFromLogicalPages(logicalPages, error) {
      const pages = (logicalPages || []).map((page, index) => {
        const pageNo = Number(page && (page.pageNo || page.sourcePageIndex) || index + 1);
        const items = pdfLogicalFallbackTextItems(page, pageNo);
        const lines = (page && page.lines && page.lines.length ? page.lines : groupPdfJsTextItemsToOverlayLines(items, pageNo))
          .filter(line => keepPdfOverlayLineNotingArtifactDrop(line, pageNo));
        const horizontalMetadataLines = pdfJsHorizontalMetadataLinesForBodyItems(items, lines, pageNo);
        return Object.assign({}, page, {
          pageNo,
          sourcePageIndex: pageNo,
          textItems: items,
          lines,
          bodyBaseLines: lines,
          horizontalMetadataLines,
          pdfJsMergedBodyLines: lines,
          textSource: 'pdf-logical-fallback'
        });
      });
      applyStablePdfJsBodyCutLineToPages(pages);
      const pdfAnnotations = pdfAnnotationsFromPages(pages);
      const warning = error && error.message
        ? 'PDF.js failed; logical fallback used: ' + error.message
        : 'PDF.js failed; logical fallback used.';
      return {
        type: 'pdf',
        text: pages.map(page => page.text || '').join('\n\f\n'),
        pages,
        pdfAnnotations,
        pdfExtractionSignals: pdfReviewSignalsFromAnnotations(pdfAnnotations),
        warning
      };
    }

    async function readPdfLogicalPagesFromBuffer(buffer, selectedPages) {
      try {
        if (typeof DecompressionStream === 'undefined') {
          // 자체 파서가 통째로 꺼진다 — pdf.js 가 CMap 을 못 읽는 문서에서는 이쪽이 유일한 본문이라 그 문서가 통째로 빈다.
          notePdfOverlayNote('logicalParserUnavailable', 'DecompressionStream');
          return [];
        }
        const streams = await extractPdfStreams(buffer);
        const pdfSource = new TextDecoder('latin1').decode(buffer);
        const decodeContext = buildPdfDecodeContext(streams, pdfSource);
        const contentStreams = filterPdfPageContentStreams(streams, pdfSource);
        const pages = await extractPdfPositionedPages(contentStreams, decodeContext, selectedPages);
        return pages.length ? attachPdfLinkAnnotationsToPages(pages, pdfSource) : [];
      } catch (error) {
        console.warn('PDF logical text extraction failed; using PDF.js overlay text.', error);
        return [];
      }
    }

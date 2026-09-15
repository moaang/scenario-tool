    function syncSourceAnnotationNotesForDocument(project, doc) {
      for (const block of (project.sourceBlocks || []).filter(b => b.sourceDocumentId === doc.id && b.sourceAnnotation && !b.sourceAnnotation.href)) {
        const source = block.sourceAnnotation;
        const originKey = [doc.id, source.trigger, source.sourceId || source.id].join('|');
        const existing = (project.revisionNotes || []).find(note => note.sourceAnnotationKey === originKey);
        if (existing) {
          existing.targetType = 'block'; existing.targetId = block.id;
          block.revisionNoteIds = uniq((block.revisionNoteIds || []).concat(existing.id));
          continue;
        }
        const note = { id: seededId('source-note', originKey), sourceAnnotationKey: originKey,
          sourceAuthor: source.author || '', sourceOriginalText: source.text || block.rawText,
          targetType: 'block', targetId: block.id, noteType: 'fact', text: source.text || block.rawText,
          linkedEntityIds: [], status: 'open', color: 'green', createdAt: nowIso() };
        project.revisionNotes.push(note); block.revisionNoteIds = uniq((block.revisionNoteIds || []).concat(note.id));
      }
    }

    function reconcileSourceBlockIdentity(blocks, project, doc) {
      const previous = (project.sourceBlocks || []).filter(b => b.sourceDocumentId === doc.id);
      const key = b => [b.type, normalizePdfCoordinateMatchText(b.rawText || b.text || '')].join('|');
      const buckets = new Map(), used = new Set(), reserved = new Set(previous.map(b => b.id));
      previous.forEach(b => { const k = key(b); if (!buckets.has(k)) buckets.set(k, []); buckets.get(k).push(b); });
      for (const block of blocks) {
        const bucket = buckets.get(key(block)) || [];
        const old = bucket.find(b => !used.has(b) && b.reviewItemId && b.reviewItemId === block.reviewItemId)
          || bucket.find(b => !used.has(b) && Number(b.pageNo) === Number(block.pageNo));
        if (old) {
          used.add(old); block.id = old.id;
          block.revisionNoteIds = (old.revisionNoteIds || []).slice();
          block.linkedEntityIds = (old.linkedEntityIds || []).slice();
          copySourceReviewGeometry(block, old);
        } else {
          const seed = block.id;
          let suffix = 0;
          while (reserved.has(block.id)) block.id = seed + '-r' + (++suffix);
        }
        reserved.add(block.id);
      }
      const remaining = new Map(previous.filter(b => !used.has(b)).map(b => [b.id, b]));
      for (const note of project.revisionNotes || []) {
        const old = note.targetType === 'block' && remaining.get(note.targetId);
        if (!old) continue;
        note.previousBlockAnchor = { blockId: old.id, text: old.rawText || old.text, pageNo: old.pageNo };
        note.targetType = 'page'; note.targetId = doc.id + '#p' + (old.pageNo || 1);
        note.sourceDocumentId = doc.id; note.pageNo = old.pageNo || 1;
      }
      return blocks;
    }

    // Confirmed imports normalize blocks, assign document identity and reinforce cast/scene
    // context here. Review previews stop before this boundary; persistence belongs to the caller.
    function finalizeImportedBlocks(parsedBlocks, doc) {
      const docId = doc && doc.id || '';
      const fileName = doc && doc.importedFileName || '';
      const blocks = normalizeImportedBlocks(parsedBlocks || []).map((block, index) => {
        const localId = block.id || ('b' + pad(index + 1, 4));
        return Object.assign(block, {
          id: docId ? docId + '-' + localId.replace(new RegExp('^' + escapeRegExp(docId) + '-'), '') : localId,
          sourceDocumentId: docId,
          sourceFileName: fileName || block.sourceFileName || ''
        });
      });
      reinforceCastBlocks(blocks);
      attachSceneContext(blocks);
      return blocks;
    }

    function normalizeImportedBlocks(blocks) {
      const normalized = [];
      (blocks || []).forEach((block, index) => {
        const id = block.id || ('b' + pad(index + 1, 4));
        const confidence = Number(block.meta && block.meta.confidence) || Number(block.confidence) || 0.55;
        const shaped = {
          id,
          sourceDocumentId: block.sourceDocumentId || 'doc-001',
          sourceFileName: block.sourceFileName || '',
          pageNo: block.pageNo || 1,
          order: block.order || block.orderOnPage || index + 1,
          globalOrder: block.globalOrder || index + 1,
          reviewItemId: block.reviewItemId || '',
          sourceAnnotation: block.sourceAnnotation ? cloneJson(block.sourceAnnotation) : null,
          overlayRect: block.overlayRect ? Object.assign({}, block.overlayRect) : null,
          pageSpans: Array.isArray(block.pageSpans) ? cloneJson(block.pageSpans) : [],
          rawText: block.rawText || block.text || '',
          text: block.text || block.normalizedText || block.rawText || '',
          normalizedText: block.normalizedText || norm(block.text || block.rawText || ''),
          inlineRuns: normalizeMarkdownRuns(block.inlineRuns || [], block.text || block.rawText || ''),
          markdownSourceLine: block.markdownSourceLine || '',
          markdownHeadingLevel: Number(block.markdownHeadingLevel) || 0,
          markdownLineIndex: Number.isInteger(block.markdownLineIndex) ? block.markdownLineIndex : -1,
          type: block.type || UNCLASSIFIED_BLOCK_TYPE,
          zone: block.zone || 'unknown',
          speaker: block.speaker || '',
          speakerNames: Array.isArray(block.speakerNames) ? block.speakerNames : splitDialogSpeakerNames(block.speaker || block.characterName || ''),
          modifier: block.modifier || '',
          modifierRaw: block.modifierRaw || '',
          dialogMode: block.dialogMode || '',
          characterName: block.characterName || block.speaker || '',
          characterGender: block.characterGender || '',
          characterAge: block.characterAge || '',
          characterProfile: Object.assign({}, block.characterProfile || {}),
          locationName: block.locationName || '',
          timeOfDay: block.timeOfDay || '',
          scenePart: block.scenePart || '',
          assignableToCut: Boolean(block.assignableToCut),
          assignDefault: block.assignDefault || 'none',
          currentScene: block.currentScene || block.locationName || '',
          linkedEntityIds: (block.linkedEntityIds || []).slice(),
          revisionNoteIds: (block.revisionNoteIds || []).slice(),
          meta: {
            confidence,
            evidence: block.meta && block.meta.evidence || [],
            warnings: block.meta && block.meta.warnings || [],
            alternatives: block.meta && block.meta.alternatives || [],
            scores: block.meta && block.meta.scores || {}
          }
        };
        normalized.push(shaped);
      });
      return normalized;
    }

    function buildSourceReviewItemsForDocument(text, doc, options) {
      // Only PDF input needs glyph cleanup; semantic classification and geometry attachment are shared.
      const isPdf = Boolean(doc && doc.draftSourceType === 'pdf');
      const pdfBodyText = isPdf ? pdfScriptBodyReviewTextFromDraftPages(doc) : '';
      const sourceText = isPdf ? cleanPdfGarbledTextLines(pdfBodyText || text) : text;
      const reviewItems = buildSourceReviewItems(sourceText, doc);
      attachSourceReviewPageSpans(reviewItems, sourceReviewPagesFromText(sourceText));
      normalizePdfInlineSpacingForReviewItems(reviewItems);
      // Annotation/metadata note items must exist BEFORE rect attachment — they used to be
      // appended after it and stayed rect-less (their rects then came from render-time
      // fallbacks that matched note text into the body stream, producing giant wrong boxes).
      appendPdfMetadataReviewItems(reviewItems, doc, options || sourceReviewOptionsForDoc(doc));
      if (doc && (doc.draftPages || []).some(page => (page.textItems || []).length)) {
        attachPdfOverlayRectsToReviewItems(reviewItems, doc.draftPages || []);
      }
      return reviewItems;
    }

    // A quoted noun continues unfinished action only when it is not an established speaker.
    // Punctuation after a genuine dialogue does not change its semantic type.
    function foldTrailingPunctuationAfterQuotedDialogItems(items) {
      if (!Array.isArray(items) || items.length < 3) return items;
      const dialogFamily = new Set(['dialog', 'monologue', 'narration']);
      const speakers = new Map();
      const cast = new Set();
      for (const item of items) {
        if (dialogFamily.has(item.type) && item.speaker) speakers.set(item.speaker, (speakers.get(item.speaker) || 0) + 1);
        if (item.type === 'character') extractCastNames(sourceReviewItemText(item)).forEach(name => cast.add(name));
      }
      for (let i = items.length - 1; i >= 2; i -= 1) {
        const cur = items[i], prev = items[i - 1], before = items[i - 2];
        if (!dialogFamily.has(prev.type) || before.type !== 'action') continue;
        if (cur.pageNo !== prev.pageNo || before.pageNo !== prev.pageNo) continue;
        if (cast.has(prev.speaker) || speakers.get(prev.speaker) > 1) continue;
        const punctuation = norm(sourceReviewItemText(cur)).trim();
        if (!/^[。、．，！？!?」』）)】〕〉》・…‥ー―─－]+$/u.test(punctuation)) continue;
        const prevText = sourceReviewItemText(prev);
        if (!/[」』）)】〕〉》”"]$/u.test(prevText.trim())) continue;
        if (sourceReviewActionEndsWithCompleteSentence(sourceReviewItemText(before))) continue;
        Object.assign(prev, { type: 'action', speaker: '', modifier: '', modifierRaw: '', dialogMode: '', dialog: '', text: prevText + punctuation });
        items.splice(i, 1);
      }
      return items;
    }

    // A wrap before a quoted noun is not a semantic boundary after both pieces are classified as action.
    function mergeWrappedSourceReviewActions(items) {
      for (let index = 1; index < items.length; index += 1) {
        const previous = items[index - 1];
        const next = items[index];
        if (previous.type !== 'action' || next.type !== 'action' || previous.pageNo !== next.pageNo) continue;
        const before = sourceReviewItemText(previous);
        const after = sourceReviewItemText(next);
        if (!before || !after || sourceReviewActionEndsWithCompleteSentence(before)) continue;
        if (isPdfTransitionMarkerLeadLine(after) || isPdfTransitionMarkerLeadLine(before)) continue;
        previous.text = before + '\n' + after;
        previous.tags = uniq((previous.tags || []).concat(next.tags || []));
        items.splice(index, 1);
        index -= 1;
      }
      renumberSourceReviewItems(items);
      return items;
    }

    function buildSourceReviewItems(text, doc) {
      const items = [];
      const pages = sourceReviewPagesFromText(text);
      const profile = createScenarioParserProfile(pages);
      const orderRef = { value: 0 };
      const logicalPages = buildScenarioLogicalPageLines(pages, profile);
      // 표지와 본문의 경계는 **문서 전체를 보고 한 번** 정한다 — 그 장과 뒤 세 장이 모두 본문답게
      // 이어지는 자리다. 예전에는 페이지마다 다시 묻고(잠금) 줄마다 또 뒤집어서, 셋이 서로를
      // 덮어쓰며 표지 한 장이 중간에 갈렸다. **확정이 하나면 잠금이라는 개념 자체가 필요 없다.**
      // 본문 분류(classifyTextToBlocks)도 같은 자를 쓴다.
      const foundReviewBodyStart = logicalPages.findIndex((page, index) => scenarioBodyStartsAtPage(logicalPages, index, profile));
      // **못 찾았다와 첫 장부터 본문이다는 다른 말이다.** 못 찾으면 0 으로 떨어뜨렸는데, 그러면
      // 「표지보다 앞」이 되는 장이 하나도 없어 표지가 통째로 본문이 된다. 세로 조판에서 대사가
      // 두 줄로 쪼개지면 본문 판정이 그 대사를 못 세어 어느 장도 본문으로 인정되지 않는다.
      // 못 찾았으면 표지가 아닌 첫 장을 경계로 삼는다. 그것마저 없으면 전부 표지다.
      const frontFallbackStart = logicalPages.findIndex((page, index) =>
        !looksLikeFrontMatterPage(page.lines || [], page.pageNo || index + 1, profile));
      const reviewBodyStartIndex = foundReviewBodyStart >= 0
        ? foundReviewBodyStart
        : (frontFallbackStart < 0 ? logicalPages.length : frontFallbackStart);
      let reviewBodyEnded = false;
      pages.forEach((page, pageIndex) => {
        const pageText = (page.lines || []).join('\n');
        const markdownItems = buildMarkdownSourceReviewItems(pageText, page.pageNo, orderRef, profile);
        if (markdownItems.length) {
          items.push(...markdownItems);
          if (markdownItems.some(item => item.type === 'end')) reviewBodyEnded = true;
          return;
        }
        const taggedItems = buildTaggedSourceReviewItems(pageText, page.pageNo, orderRef, profile);
        if (taggedItems.length) {
          items.push(...taggedItems);
          if (taggedItems.some(item => item.type === 'end')) reviewBodyEnded = true;
          return;
        }
        const reviewFrontOpen = pageIndex < reviewBodyStartIndex;
        const logicalLines = (logicalPages[pageIndex] && logicalPages[pageIndex].lines) || [];
        const lines = sourceReviewLinesForPage(page, logicalLines, reviewFrontOpen, profile);
        let buffer = [];
        let bufferType = 'text';
        let bufferSection = '';
        let inCast = false;
        const pageBodyLikely = pageHasBodySignal(lines, profile);
        let bodyStarted = !reviewFrontOpen;
        const candidateRuns = detectFrontCharacterLineRuns(lines);
        const candidateStarts = new Map(candidateRuns.map(run => [run.start, run]));
        let bufferFrontKind = '';
        const flushText = () => {
          if (!buffer.length) return;
          const item = pushSourceReviewItem(items, bufferType, page.pageNo, buffer.join('\n'), orderRef);
          if (item && reviewFrontOpen) markSourceReviewSection([item], 'FRONT');
          if (item && bufferSection) markSourceReviewSection([item], bufferSection);
          buffer = [];
          bufferSection = '';
          bufferFrontKind = '';
        };
        const appendBuffered = (type, raw, section) => {
          const nextSection = String(section || '').trim().toLowerCase();
          // 표지에서는 **줄마다 종류가 바뀐다** — 작품명 · 화수 · 원고 차수 · 날짜 · 각본가.
          // 종류가 바뀌는데도 이어 붙이면 그 다섯이 한 덩어리가 되고, 뒤에서 세분기가 첫 줄만 보고
          // 판정한다(실제로 그래서 표지 전체가 한 종류로 나왔다). 같은 종류가 이어지는 것은
          // 한 문장이 두 줄로 접힌 경우이므로 그대로 붙인다.
          const nextFrontKind = reviewFrontOpen && type === 'text' ? classifyFront(norm(raw), 0).type : '';
          if (buffer.length && (bufferType !== type || bufferSection !== nextSection || bufferFrontKind !== nextFrontKind)) flushText();
          bufferType = type;
          bufferSection = nextSection;
          bufferFrontKind = nextFrontKind;
          buffer.push(raw);
        };
        for (let index = 0; index < lines.length; index += 1) {
          const raw = lines[index];
          if (reviewBodyEnded) {
            if (isEndMarkerLine(raw)) {
              flushText();
              pushSourceReviewItem(items, 'end', page.pageNo, raw, orderRef);
              continue;
            }
            appendBuffered('text', raw, 'appendix');
            continue;
          }
          const candidateRun = candidateStarts.get(index);
          if (reviewFrontOpen && candidateRun) {
            flushText();
            pushFrontCastLines(items, page.pageNo, candidateRun.lines, orderRef);
            index = candidateRun.end;
            continue;
          }
          if (reviewFrontOpen) {
            if (isCastSectionLine(raw)) {
              appendBuffered('text', raw);
              inCast = true;
              continue;
            }
            if (inCast && isExplicitCastListNameLine(raw, profile)) {
              flushText();
              pushFrontCastItem(items, page.pageNo, raw, orderRef);
              continue;
            }
            if (inCast && isCastGroupMarkerLine(raw)) {
              flushText();
              pushFrontSourceReviewItem(items, 'text', page.pageNo, raw, orderRef);
              continue;
            }
            if (isLikelyCastEntryLine(raw, profile) && !isCastGroupMarkerLine(raw)) {
              flushText();
              const collected = collectCastEntryLines(lines, index, profile, { frontContinuation: reviewFrontOpen });
              index = collected.endIndex;
              pushFrontCastItem(items, page.pageNo, collected.lines.join('\n'), orderRef);
              continue;
            }
            appendBuffered('text', raw);
            continue;
          }
          if (isCastSectionLine(raw)) {
            appendBuffered('text', raw);
            inCast = true;
            continue;
          }
          if (inCast && isCastGroupMarkerLine(raw)) {
            flushText();
            pushSourceReviewItem(items, 'text', page.pageNo, raw, orderRef);
            continue;
          }
          if (inCast && isExplicitCastListNameLine(raw, profile)) {
            flushText();
            pushSourceReviewItem(items, 'character', page.pageNo, raw, orderRef);
            continue;
          }
          if (inCast && isLikelyCastEntryLine(raw, profile) && !isCastGroupMarkerLine(raw)) {
            flushText();
            const collected = collectCastEntryLines(lines, index, profile, {});
            index = collected.endIndex;
            pushSourceReviewItem(items, 'character', page.pageNo, collected.lines.join('\n'), orderRef);
            continue;
          }
          if (parseDialog(raw, profile)) {
            if (buffer.length && bufferType === 'text' && pageBodyLikely) bufferType = 'action';
            flushText();
            pushSourceReviewItem(items, 'dialog', page.pageNo, raw, orderRef);
            bodyStarted = true;
            inCast = false;
            continue;
          }
          if (lineHasDialogStartSignal(raw, profile)) {
            if (buffer.length && bufferType === 'text' && pageBodyLikely) bufferType = 'action';
            flushText();
            const collected = collectSourceReviewOpenDialogLines(lines, index, profile);
            pushSourceReviewItem(items, 'dialog', page.pageNo, collected.text, orderRef);
            index = collected.endIndex;
            bodyStarted = true;
            inCast = false;
            continue;
          }
          if (isEndMarkerLine(raw)) {
            flushText();
            pushSourceReviewItem(items, 'end', page.pageNo, raw, orderRef);
            bodyStarted = false;
            inCast = false;
            reviewBodyEnded = true;
            continue;
          }
          if (isSectionMarkerLine(raw)) {
            flushText();
            pushSourceReviewItem(items, 'scene', page.pageNo, raw, orderRef);
            bodyStarted = true;
            inCast = false;
            continue;
          }
          if (isStandaloneStructureMarkerLine(raw)) {
            flushText();
            pushSourceReviewItem(items, 'text', page.pageNo, raw, orderRef);
            bodyStarted = true;
            inCast = false;
            continue;
          }
          if (isSceneHeading(raw, profile)) {
            flushText();
            pushSourceReviewItem(items, 'scene', page.pageNo, raw, orderRef);
            bodyStarted = true;
            inCast = false;
            continue;
          }
          if (isStrongBodyStart(raw, profile) && !isCastSectionLine(raw)) {
            flushText();
            pushSourceReviewItem(items, 'scene', page.pageNo, raw, orderRef);
            bodyStarted = true;
            inCast = false;
            continue;
          }
          if (bodyStarted && (isBodyTextInsertLine(raw) || startsWithNonDominantSceneMarker(raw, profile))) {
            flushText();
            const textLines = [raw];
            while (index + 1 < lines.length && isBodyTextInsertContinuationLine(lines[index + 1], profile)) {
              textLines.push(lines[index + 1]);
              index += 1;
            }
            pushSourceReviewItem(items, 'text', page.pageNo, textLines.join('\n'), orderRef);
            continue;
          }
          if (isPdfTransitionMarkerLeadLine(raw)) {
            // A ××× line marks a shot/beat transition WITHIN a scene (cf. classifyBody's
            // ^××× transition rule). It's often extracted glued to the FOLLOWING sentence
            // rather than on its own line, so match on the LEADING marker (not anchored at
            // the end) and flush first — that ends whatever action beat came before it, and
            // pushing this line (marker ± its trailing text) as its own item, bypassing the
            // buffer, keeps the NEXT line from merging backward into it too. Net effect: the
            // action beats on either side of ××× become distinct items/overlay boxes instead
            // of one merged item spanning the wrong rect.
            flushText();
            const hasTrailingContent = Boolean(norm(raw).replace(new RegExp('^' + PDF_TRANSITION_MARK_CLASS + '{2,}\\s*'), '').trim());
            pushSourceReviewItem(items, hasTrailingContent && bodyStarted ? 'action' : 'text', page.pageNo, raw, orderRef);
            continue;
          }
          appendBuffered(bodyStarted ? 'action' : 'text', raw);
        }
        flushText();
      });
      mergeSourceReviewActionPageContinuations(items, pages, profile);
      foldTrailingPunctuationAfterQuotedDialogItems(items);
      mergeWrappedSourceReviewActions(items);
      if (!items.length) pushSourceReviewItem(items, 'text', 1, doc && doc.originalText || text || '', orderRef);
      return items;
    }

    // Join incomplete action text across adjacent physical pages for every source format.
    // The common span/geometry pass subsequently anchors the merged text on each source page.
    function mergeSourceReviewActionPageContinuations(items, pages, profile) {
      if (!Array.isArray(items) || items.length < 2 || !Array.isArray(pages) || pages.length < 2) return items || [];
      const pageNos = pages.map(page => Number(page && page.pageNo || 0)).filter(pageNo => pageNo > 0);
      let changed = false;
      for (let index = 0; index < pageNos.length - 1; index += 1) {
        const pageNo = pageNos[index];
        const nextPageNo = pageNos[index + 1];
        if (nextPageNo !== pageNo + 1) continue;
        const previous = lastBodySourceReviewItemOnPage(items, pageNo);
        const next = firstBodySourceReviewItemOnPage(items, nextPageNo);
        if (!previous || !next || previous.index >= next.index) continue;
        if (!shouldMergeSourceReviewActionPageContinuation(previous.item, next.item, profile)) continue;
        mergeSourceReviewActionItemsAcrossPages(items, previous.index, next.index);
        changed = true;
      }
      if (changed) renumberSourceReviewItems(items);
      return items;
    }

    function attachSourceReviewPageSpans(items, pages) {
      const sourcePages = (pages || []).map((page) => {
        const lines = page && page.lines || [];
        const text = lines.join('\n');
        const pageSpanLineKeys = lines.map(line => normalizePdfCompactTextNfkc(line)).filter(Boolean);
        return {
          pageNo: Number(page && page.pageNo || 1),
          text,
          pageSpanKey: normalizePdfCompactTextNfkc(text),
          pageSpanLineKeys
        };
      }).filter(page => page.pageNo && (page.pageSpanKey || page.pageSpanLineKeys.length));
      if (!sourcePages.length) return items || [];
      (items || []).forEach((item) => {
        if (!item || isSourceReviewPageArtifactItem(item) || isSourceReviewPdfColorAnnotationItem(item)) return;
        const spans = sourceReviewPageSpansForItem(item, sourcePages);
        item.pageSpans = spans.length ? spans : [fallbackPageSpanForItem(item)];
      });
      return items || [];
    }

    function carrySourceReviewItemGeometry(items, previousItems) {
      const nextItems = items || [];
      const previous = previousItems || [];
      if (!nextItems.length || !previous.length) return nextItems;
      const previousById = new Map();
      const previousByText = new Map();
      previous.forEach((item) => {
        if (!item) return;
        if (item.id) previousById.set(item.id, item);
        const key = sourceReviewGeometryMatchKey(item);
        if (!key) return;
        if (!previousByText.has(key)) previousByText.set(key, []);
        previousByText.get(key).push(item);
      });
      const used = new Set();
      nextItems.forEach((item) => {
        if (!item) return;
        let source = previousById.get(item.id || '');
        if (used.has(source)) source = null;
        if (!source) {
          // id 는 쪽·순번·본문 해시로 만들므로(makeSourceReviewItem) 앞 줄 하나만 고쳐도 뒤가 전부
          // 어긋난다 — 그래서 본문 버킷으로 떨어진다. 양쪽 목록이 다 문서 순서라, 같은 키가 여럿이면
          // k번째끼리 짝지어 상대 순서를 지킨다(shift 가 그 짝짓기다).
          const bucket = previousByText.get(sourceReviewGeometryMatchKey(item)) || [];
          while (bucket.length && used.has(bucket[0])) bucket.shift();
          source = bucket.shift() || null;
        }
        if (!source) return;
        used.add(source);
        copySourceReviewGeometry(item, source);
      });
      return nextItems;
    }

    function sourceReviewGeometryMatchKey(item) {
      if (!item) return '';
      return [
        Number(item.pageNo || 1),
        item.type || '',
        normalizePdfCoordinateMatchText(sourceReviewItemText(item))
      ].join('|');
    }

    function copySourceReviewGeometry(target, source) {
      if (!target || !source) return;
      if (Number(target.pageNo || 1) === Number(source.pageNo || 1) && source.overlayRect
          && (!target.overlayRect || source.overlayRect.manual && !target.overlayRect.manual)) target.overlayRect = cloneJson(source.overlayRect);
      if (!Array.isArray(target.pageSpans) || !target.pageSpans.length) {
        target.pageSpans = cloneJson(source.pageSpans || []);
        return;
      }
      target.pageSpans.forEach(span => {
        if (!span || span.overlayRect && span.overlayRect.manual) return;
        const matched = (source.pageSpans || []).find(old => Number(old.pageNo) === Number(span.pageNo)
          && normalizePdfCoordinateMatchText(old.text || '') === normalizePdfCoordinateMatchText(span.text || ''));
        if (matched && matched.overlayRect && (!span.overlayRect || matched.overlayRect.manual)) span.overlayRect = cloneJson(matched.overlayRect);
      });
    }

    function sourceReviewItemsToBlocks(items, doc) {
      const blocks = [];
      let globalOrder = 0;
      let currentScene = '';
      let currentLocationName = '';
      let currentTimeOfDay = '';
      let currentScenePart = '';
      // 프로파일은 **임포트 원문**으로 세운다. 인용부호·씬 기호·번호 씬 형식을 문서 전체의 절대
      // 개수로 판정하므로(후보 3건 이상 등) 편집으로 줄어든 본문에서 다시 세우면 판정이 뒤집힌다.
      // originalText 는 저장 때 편집본으로 덮이므로 뒷자리 폴백이다.
      const conversionProfile = createScenarioParserProfile(sourceReviewPagesFromText(doc && (doc.rawImportedText || doc.originalText) || ''));
      const bodyReviewPages = new Set((items || []).filter(isBodyReviewItem).map(item => Number(item.pageNo || 0)));
      const pushBlock = (item, rawText, classification) => {
        let raw = String(rawText || '').trim();
        // 본문이 빈 블록이 목록에 들어가는 것을 막는다(블록은 text 로 그려지고 매칭된다).
        // 실측: 임포트 경로 18개 문서 8560 항목에서 여기 걸린 항목은 0(항목 수 = 블록 수).
        if (!raw) return;
        globalOrder += 1;
        const c = Object.assign({
          type: 'unknown',
          zone: MANUAL_REVIEW_ZONE,
          text: raw,
          assignableToCut: false,
          assignDefault: 'none',
          confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.default,
          evidence: ['source review confirmed'],
          warnings: []
        }, classification || {});
        if (c.type === 'scene_heading') {
          raw = normalizeSceneHeadingText(raw);
          c.text = raw;
          const meta = parseSceneMeta(raw, currentLocationName);
          c.locationName = meta.locationName || c.locationName || '';
          c.timeOfDay = meta.timeOfDay || c.timeOfDay || '';
          c.scenePart = meta.scenePart || c.scenePart || '';
          if (c.locationName) currentLocationName = c.locationName;
          else if (c.scenePart) currentLocationName = '';
          if (c.timeOfDay) currentTimeOfDay = c.timeOfDay;
          if (c.scenePart) currentScenePart = c.scenePart;
          currentScene = raw;
        }
        if (c.type === 'section_marker' && c.scenePart) {
          currentScenePart = c.scenePart;
          currentLocationName = '';
        }
        if (c.zone === 'body') {
          c.currentScene = c.currentScene || currentScene;
          c.locationName = c.locationName || currentLocationName;
          c.timeOfDay = c.timeOfDay || currentTimeOfDay;
          c.scenePart = c.scenePart || currentScenePart;
        }
        const block = makeBlock('b' + pad(globalOrder, 4), raw, c, item.pageNo || 1, item.order || globalOrder, globalOrder);
        if (item.overlayRect) block.overlayRect = Object.assign({}, item.overlayRect);
        if (item.pageSpans && item.pageSpans.length) block.pageSpans = cloneJson(item.pageSpans);
        block.reviewItemId = item.id || '';
        if (item.sourceAnnotation) block.sourceAnnotation = cloneJson(item.sourceAnnotation);
        blocks.push(block);
        return block;
      };
      (items || []).forEach((item) => {
        if (!item) return;
        if (isSourceReviewPageArtifactItem(item)) return;
        if (isSourceReviewPdfColorAnnotationItem(item)) {
          pushBlock(item, sourceReviewItemText(item), {
            type: 'pdf_note',
            zone: 'page_metadata',
            assignableToCut: false,
            assignDefault: 'none',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.pdf_note,
            evidence: ['source review pdf note']
          });
          return;
        }
        if (item.type === 'character') {
          const raw = sourceReviewItemText(item);
          const profile = parseCharacterProfileFromCastText(raw, item.name);
          pushBlock(item, raw, {
            type: 'cast_entry',
            zone: 'cast_list',
            text: raw,
            characterName: item.name || profile.name || extractCastNames(raw)[0] || '',
            characterGender: item.gender || profile.gender || '',
            characterAge: item.age || profile.age || '',
            characterProfile: Object.assign({}, profile, item.profile || {}),
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.cast_entry,
            evidence: item.section === 'front' ? ['front cast'] : ['source review cast']
          });
          return;
        }
        if (item.type === 'dialog') {
          const raw = sourceReviewItemText(item);
          const dialogBlockType = item.dialogMode === 'monologue' ? 'monologue' : (item.dialogMode === 'narration' ? 'narration' : 'dialog');
          pushBlock(item, raw, {
            type: dialogBlockType,
            zone: 'body',
            text: item.dialog || dialogTextWithQuotes(item),
            speaker: item.speaker || '',
            speakerNames: splitDialogSpeakerNames(item.speaker || ''),
            modifier: item.modifier || '',
            modifierRaw: item.modifierRaw || '',
            dialogMode: item.dialogMode || '',
            openQuote: item.openQuote || '',
            closeQuote: item.closeQuote || '',
            characterName: item.speaker || '',
            assignableToCut: true,
            assignDefault: 'dialog',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.dialog,
            evidence: ['source review dialog']
          });
          return;
        }
        if (item.type === 'scene') {
          const sceneRaw = sourceReviewItemText(item);
          if (isPartMarkerText(sceneRaw, conversionProfile)) {
            // Part label (アバン / Aパート): a part-membership marker, not a scene.
            pushBlock(item, sceneRaw, {
              type: 'section_marker',
              zone: 'body',
              assignableToCut: true,
              assignDefault: 'action',
              scenePart: partMarkerLabel(sceneRaw),
              confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.section_marker,
              evidence: ['source review part marker']
            });
            return;
          }
          pushBlock(item, sceneRaw, {
            type: 'scene_heading',
            zone: 'body',
            assignableToCut: true,
            assignDefault: 'action',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.scene_heading,
            evidence: ['source review scene']
          });
          return;
        }
        if (item.type === 'action') {
          pushBlock(item, item.text || sourceReviewItemText(item), {
            type: 'action',
            zone: 'body',
            assignableToCut: true,
            assignDefault: 'action',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.action,
            evidence: ['source review action']
          });
          return;
        }
        if (item.type === 'end') {
          pushBlock(item, sourceReviewItemText(item), {
            type: 'end_marker',
            zone: 'body_end',
            assignableToCut: false,
            assignDefault: 'none',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.end_marker,
            evidence: ['source review end marker']
          });
          currentScene = '';
          currentLocationName = '';
          currentTimeOfDay = '';
          currentScenePart = '';
          return;
        }
        if (item.type === 'note') {
          pushBlock(item, sourceReviewItemText(item), {
            type: 'pdf_note',
            zone: 'page_metadata',
            assignableToCut: false,
            assignDefault: 'none',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.pdf_note,
            evidence: ['source review note']
          });
          return;
        }
        if (item.type === 'link') {
          pushBlock(item, sourceReviewItemText(item), {
            type: 'pdf_link',
            zone: 'page_metadata',
            assignableToCut: false,
            assignDefault: 'none',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.pdf_link,
            evidence: ['source review link']
          });
          return;
        }
        const reviewText = sourceReviewItemText(item);
        if (isAppendixSourceReviewTextItem(item)) {
          pushBlock(item, reviewText, {
            type: 'appendix_text',
            zone: 'appendix',
            assignableToCut: false,
            assignDefault: 'none',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.appendix_text,
            evidence: ['source review appendix text']
          });
          return;
        }
        const pageHasBodyReview = bodyReviewPages.has(Number(item.pageNo || 0));
        const textIsFront = isFrontSourceReviewTextItem(item);
        const textInScene = !textIsFront && (item.section === 'scene'
          || (item.tags || []).includes('section:scene')
          || pageHasBodyReview
          || isBodyTextInsertLine(reviewText)
          || startsWithNonDominantSceneMarker(reviewText, conversionProfile));
        if (textInScene) {
          pushBlock(item, reviewText, {
            type: 'body_text',
            zone: 'body',
            assignableToCut: false,
            assignDefault: 'none',
            confidence: SOURCE_REVIEW_BLOCK_CONFIDENCE.body_text,
            evidence: ['source review body text']
          });
          return;
        }
        // 표지 줄은 작품명·화수·원고 차수·날짜·각본가로 갈린다. **그 판정은 이미 classifyFront 에
        // 있다** — 여기서 다시 세우면 두 경로가 표지를 다르게 읽어 화면과 데이터가 갈린다.
        // (실제로 갈려 있었다: 텍스트 임포트는 다섯 종류로 나누는데 PDF 는 전부 front_summary 였다.)
        const frontClass = classifyFront(reviewText, frontScore(reviewText, Number(item.pageNo || 1), Number(item.order || 0), conversionProfile));
        // classifyFront 가 「이건 등장인물 목록의 머리다」라고 판정하면 그 zone 을 쓴다.
        // 예전에는 여기서 무조건 덮어써서 **타입만 살고 zone 은 조용히 탈락했다** —
        // 코드는 cast_list 라고 적혀 있는데 실제 블록은 front_matter 로 나왔다.
        const frontZone = frontClass.zone === 'cast_list'
          ? 'cast_list'
          : (textIsFront ? 'front_matter' : 'manual_text');
        pushBlock(item, reviewText, Object.assign({}, frontClass, {
          zone: frontZone,
          assignableToCut: false,
          assignDefault: 'none',
          evidence: [textIsFront ? 'source review front text' : 'source review text']
        }));
      });
      return blocks;
    }

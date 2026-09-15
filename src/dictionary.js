    function defaultUserDictionary() {
      return { presets: [{ id: 'sf-terms', name: 'SF', enabled: true,
        entries: ['AI', 'ドローン', 'ハッカー', 'ハッカールーム', 'ネットワーク', 'セーフハウス', 'サイボーグ']
          .map(term => ({ term, type: 'technology' })) }] };
    }

    function newDictionaryPresetId() {
      return seededId('preset', nowIso() + Math.random());
    }

    function normalizeUserDictionary(value) {
      if (!value || !Array.isArray(value.presets)) return defaultUserDictionary();
      const ids = new Set();
      const presets = value.presets.filter(preset => preset && typeof preset === 'object').map(preset => {
        let id = typeof preset.id === 'string' ? preset.id : '';
        if (!id || ids.has(id)) id = newDictionaryPresetId();
        ids.add(id);
        const seen = new Set();
        const entries = (Array.isArray(preset.entries) ? preset.entries : []).filter(entry => {
          if (!entry || typeof entry.term !== 'string' || !entry.term.trim() || !ENTITY_TYPES.includes(entry.type)) return false;
          const key = entry.type + '|' + entry.term.trim();
          if (seen.has(key)) return false;
          seen.add(key); return true;
        }).map(entry => ({ term: entry.term.trim(), type: entry.type }));
        return { id, name: String(preset.name || '').trim() || 'Preset', enabled: preset.enabled !== false, entries };
      });
      return { presets };
    }

    function activeUserDictionaryEntries(dictionary) {
      const entries = [], seen = new Set();
      for (const preset of normalizeUserDictionary(dictionary).presets) {
        if (!preset.enabled) continue;
        for (const entry of preset.entries) {
          const key = entry.type + '|' + entry.term;
          if (!seen.has(key)) { seen.add(key); entries.push(entry); }
        }
      }
      return entries;
    }

    function parseUserDictionaryFile(text) {
      const value = JSON.parse(text);
      if (!value || value.format !== 'scenario-user-dictionary-v1' || !Array.isArray(value.presets)) throw Error(t('dictionaryInvalidFile'));
      for (const preset of value.presets) {
        if (!preset || typeof preset.name !== 'string' || typeof preset.enabled !== 'boolean' || !Array.isArray(preset.entries)) throw Error(t('dictionaryInvalidFile'));
        for (const entry of preset.entries) {
          if (!entry || typeof entry.term !== 'string' || !entry.term.trim() || entry.term.length > 80 || !ENTITY_TYPES.includes(entry.type)) throw Error(t('dictionaryInvalidFile'));
        }
      }
      return normalizeUserDictionary({ presets: value.presets.map(preset => Object.assign({}, preset, { id: newDictionaryPresetId() })) });
    }

    function serializeUserDictionary(dictionary) {
      return JSON.stringify(Object.assign({ format: 'scenario-user-dictionary-v1' }, normalizeUserDictionary(dictionary)), null, 2);
    }

    function openAppSettings() {
      closeModal();
      const draft = cloneJson(normalizeUserDictionary(state.settings.userDictionary));
      let selectedId = draft.presets[0] && draft.presets[0].id;
      const modal = $('editModal');
      modal.querySelector('.modal').classList.add('settings-modal');
      $('modalTitle').textContent = t('appSettings');
      modal.classList.add('active');
      const render = () => {
        const selected = draft.presets.find(preset => preset.id === selectedId);
        $('modalBody').innerHTML = `
          <div class="dictionary-settings">
            <div role="tablist" aria-label="${esc(t('appSettings'))}"><button type="button" role="tab" aria-selected="true" aria-controls="dictionaryPanel" id="dictionaryTab" class="primary">${esc(t('userDictionary'))}</button></div>
            <p class="dictionary-help">${esc(t('dictionaryHelp'))}</p>
            <div id="dictionaryPanel" role="tabpanel" aria-labelledby="dictionaryTab" class="dictionary-layout">
              <aside class="dictionary-presets">
                <div class="dictionary-preset-list">${draft.presets.map(preset => `<div class="dictionary-preset-row">
                  <input type="checkbox" data-preset-enabled="${esc(preset.id)}" aria-label="${esc(t('dictionaryEnabled') + ': ' + preset.name)}" ${preset.enabled ? 'checked' : ''}>
                  <button type="button" data-preset-select="${esc(preset.id)}" aria-pressed="${preset.id === selectedId}" class="${preset.id === selectedId ? 'primary' : ''}">${esc(preset.name)}</button>
                </div>`).join('')}</div>
                <button type="button" data-dictionary-action="add-preset">${esc(t('dictionaryAddPreset'))}</button>
                <button type="button" data-dictionary-action="import">${esc(t('dictionaryImport'))}</button>
                <button type="button" data-dictionary-action="export-all">${esc(t('dictionaryExportAll'))}</button>
                <input type="file" data-dictionary-file accept=".json,application/json" hidden>
              </aside>
              <section class="dictionary-editor">${selected ? `
                <label>${esc(t('dictionaryPresetName'))}<input type="text" data-preset-name value="${esc(selected.name)}" maxlength="80"></label>
                <div class="dictionary-tools">
                  <button type="button" data-dictionary-action="export">${esc(t('dictionaryExport'))}</button>
                  <button type="button" class="danger" data-dictionary-action="delete-preset">${esc(t('dictionaryDeletePreset'))}</button>
                </div>
                <div class="dictionary-entry-list">${selected.entries.map((entry, index) => `<div class="dictionary-entry" data-entry-index="${index}">
                  <input type="text" data-entry-term value="${esc(entry.term)}" maxlength="80" aria-label="${esc(t('dictionaryTerm'))}" placeholder="${esc(t('dictionaryTerm'))}">
                  <select data-entry-type aria-label="${esc(t('dictionaryType'))}">${ENTITY_TYPES.map(type => `<option value="${type}" ${type === entry.type ? 'selected' : ''}>${esc(typeLabel(type))}</option>`).join('')}</select>
                  <button type="button" data-dictionary-action="delete-entry" aria-label="${esc(t('dictionaryDeleteEntry'))}">×</button>
                </div>`).join('')}</div>
                <button type="button" data-dictionary-action="add-entry">${esc(t('dictionaryAddEntry'))}</button>
              ` : `<p>${esc(t('dictionaryEmpty'))}</p>`}</section>
            </div>
            <p class="dictionary-status" role="status"></p>
            <div class="dictionary-footer"><button type="button" class="primary" data-dictionary-action="save">${esc(t('dictionarySave'))}</button></div>
          </div>`;
        const root = $('modalBody').querySelector('.dictionary-settings');
        root.addEventListener('input', event => {
          const target = event.target;
          if (target.matches('[data-preset-name]')) {
            selected.name = target.value;
            const button = Array.from(root.querySelectorAll('[data-preset-select]')).find(button => button.dataset.presetSelect === selected.id);
            button.textContent = target.value;
          }
          const row = target.closest('[data-entry-index]');
          if (row && selected) {
            const entry = selected.entries[Number(row.dataset.entryIndex)];
            if (target.matches('[data-entry-term]')) entry.term = target.value;
            if (target.matches('[data-entry-type]')) entry.type = target.value;
          }
          if (target.matches('[data-preset-enabled]')) draft.presets.find(p => p.id === target.dataset.presetEnabled).enabled = target.checked;
        });
        root.querySelector('[data-dictionary-file]').addEventListener('change', async event => {
          const file = event.target.files[0];
          if (!file) return;
          try {
            const imported = parseUserDictionaryFile(await file.text());
            draft.presets.push(...imported.presets);
            if (imported.presets.length) selectedId = imported.presets[0].id;
            render();
          } catch (error) {
            root.querySelector('.dictionary-status').textContent = t('dictionaryImportFailed', { message: error.message });
            event.target.value = '';
          }
        });
        root.addEventListener('click', async event => {
          const presetButton = event.target.closest('[data-preset-select]');
          if (presetButton) { selectedId = presetButton.dataset.presetSelect; render(); return; }
          const button = event.target.closest('[data-dictionary-action]');
          if (!button) return;
          switch (button.dataset.dictionaryAction) {
            case 'add-preset': {
              const preset = { id: newDictionaryPresetId(), name: t('dictionaryNewPreset'), enabled: true, entries: [] };
              draft.presets.push(preset); selectedId = preset.id; render();
              $('modalBody').querySelector('[data-preset-name]').select(); return;
            }
            case 'delete-preset':
              draft.presets.splice(draft.presets.indexOf(selected), 1);
              selectedId = draft.presets[0] && draft.presets[0].id; render(); return;
            case 'add-entry':
              selected.entries.push({ term: '', type: 'technology' }); render();
              Array.from($('modalBody').querySelectorAll('[data-entry-term]')).pop().focus(); return;
            case 'delete-entry':
              selected.entries.splice(Number(button.closest('[data-entry-index]').dataset.entryIndex), 1); render(); return;
            case 'import': root.querySelector('[data-dictionary-file]').click(); return;
            case 'export':
            case 'export-all':
              downloadBlob(serializeUserDictionary(button.dataset.dictionaryAction === 'export' ? { presets: [selected] } : draft), 'scenario-user-dictionary.json', 'application/json;charset=utf-8'); return;
            case 'save': {
              button.disabled = true;
              const previous = state.settings.userDictionary;
              state.settings.userDictionary = normalizeUserDictionary(draft);
              try {
                if (!await putSettingsAutosaveSnapshot('user_dictionary')) throw Error(t('dictionaryStorageUnavailable'));
                lastAutosaveAt = Date.now(); updateSaveStatus(); closeModal(); toast(t('dictionarySaved'));
              } catch (error) {
                state.settings.userDictionary = previous;
                root.querySelector('.dictionary-status').textContent = t('dictionarySaveFailed', { message: error.message });
                button.disabled = false;
              }
            }
          }
        });
      };
      render();
    }

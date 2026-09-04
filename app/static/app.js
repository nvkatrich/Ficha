(() => {
  const byId = (id) => document.getElementById(id);
  const templateSelect = byId('templateSelect');
  const dealId = byId('dealId');
  const previewPanel = byId('previewPanel');
  const previewFields = byId('previewFields');
  const previewTitle = byId('previewTitle');
  const previewState = byId('previewState');
  const productsPreview = byId('productsPreview');
  const specificationSelect = byId('specificationSelect');
  const loadSpecificationsButton = byId('loadSpecifications');

  function syncOutputFormats() {
    const option = templateSelect.options[templateSelect.selectedIndex];
    const isPptx = option?.dataset.kind === 'pptx';
    document.querySelector('.format-pptx').classList.toggle('hidden', !isPptx);
    document.querySelector('.format-docx').classList.toggle('hidden', isPptx);
    const selected = document.querySelector('input[name="output_format"]:checked');
    if (isPptx && selected?.value === 'docx') document.querySelector('input[value="pptx"]').checked = true;
    if (!isPptx && selected?.value === 'pptx') document.querySelector('input[value="docx"]').checked = true;
  }

  function toggleAuthFields() {
    const isOAuth = document.querySelector('input[name="bitrix_auth_type"]:checked')?.value === 'oauth';
    byId('webhookFields').classList.toggle('hidden', isOAuth);
    byId('oauthFields').classList.toggle('hidden', !isOAuth);
  }
  document.querySelectorAll('input[name="bitrix_auth_type"]').forEach((input) => input.addEventListener('change', toggleAuthFields));
  toggleAuthFields();

  function clearNode(node) { while (node.firstChild) node.removeChild(node.firstChild); }
  function create(tag, text, className) {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    if (className) node.className = className;
    return node;
  }

  function renderProducts(products) {
    clearNode(productsPreview);
    if (!products.length) return;
    const caption = create('b', `Товары из сделки (${products.length})`);
    productsPreview.appendChild(caption);
    const table = document.createElement('table');
    const head = document.createElement('thead');
    const row = document.createElement('tr');
    ['№', 'Наименование', 'Кол-во', 'Цена'].forEach((value) => row.appendChild(create('th', value)));
    head.appendChild(row); table.appendChild(head);
    const body = document.createElement('tbody');
    products.forEach((product) => {
      const tr = document.createElement('tr');
      [product.number, product.name, `${product.quantity} ${product.unit || ''}`, product.price].forEach((value) => tr.appendChild(create('td', String(value ?? ''))));
      body.appendChild(tr);
    });
    table.appendChild(body); productsPreview.appendChild(table);
  }

  async function loadSpecifications() {
    const id = dealId.value.trim();
    if (!id) { loadSpecificationsButton.textContent = 'Сначала укажите ID сделки'; return; }
    const initial = loadSpecificationsButton.textContent;
    loadSpecificationsButton.textContent = 'Ищем файлы в комментариях…';
    try {
      const response = await fetch(`/api/deals/${encodeURIComponent(id)}/specifications`);
      const result = await response.json();
      if (!response.ok || !result.ok) throw new Error(result.message || 'Не удалось получить комментарии.');
      while (specificationSelect.options.length > 1) specificationSelect.remove(1);
      (result.files || []).forEach((file) => {
        const extension = (file.name || '').split('.').pop().toUpperCase();
        if (!['XLSX', 'XLS', 'XLSM', 'CSV', 'DOCX', 'PPTX', 'PDF'].includes(extension)) return;
        const option = create('option', `${file.name} · комментарий ${file.created || ''}`);
        option.value = file.key; specificationSelect.appendChild(option);
      });
      loadSpecificationsButton.textContent = result.files?.length ? `Найдено файлов: ${result.files.length}` : 'Файлы в комментариях не найдены';
      setTimeout(() => { loadSpecificationsButton.textContent = initial; }, 3500);
    } catch (error) {
      loadSpecificationsButton.textContent = error.message || 'Ошибка поиска';
      setTimeout(() => { loadSpecificationsButton.textContent = initial; }, 3500);
    }
  }

  function renderFields(fields) {
    clearNode(previewFields);
    fields.forEach((field) => {
      const box = create('div', undefined, 'preview-field');
      const label = document.createElement('label');
      const code = create('code', `{{${field.key}}}`);
      const input = document.createElement('input');
      input.name = 'override_value'; input.value = field.value || ''; input.form = 'generatorForm';
      input.dataset.overrideFor = field.key;
      const hidden = document.createElement('input');
      hidden.type = 'hidden'; hidden.name = 'override_key'; hidden.value = field.key; hidden.form = 'generatorForm';
      label.appendChild(code); label.appendChild(input); box.appendChild(label); box.appendChild(hidden);
      previewFields.appendChild(box);
    });
  }

  async function previewDeal() {
    const templateId = templateSelect.value;
    const id = dealId.value.trim();
    if (!templateId || !id) {
      previewPanel.classList.remove('hidden'); previewTitle.textContent = 'Нужны шаблон и ID сделки'; previewState.textContent = 'Заполните поля выше'; return;
    }
    previewPanel.classList.remove('hidden');
    previewTitle.textContent = 'Запрашиваем данные…'; previewState.textContent = 'Bitrix24'; clearNode(previewFields); clearNode(productsPreview);
    try {
      const specKey = specificationSelect.value;
      const response = await fetch(`/api/deals/${encodeURIComponent(id)}/preview?template_id=${encodeURIComponent(templateId)}&specification_key=${encodeURIComponent(specKey)}`);
      const result = await response.json();
      if (!response.ok || !result.ok) throw new Error(result.message || 'Не удалось загрузить сделку.');
      previewTitle.textContent = result.deal_title || `Сделка #${id}`;
      previewState.textContent = result.specification?.name ? `Источник: ${result.specification.name}` : 'Источник: товарные позиции CRM';
      renderFields(result.fields); renderProducts(result.products || []);
      if (result.specification?.warnings?.length) previewState.textContent += ` · предупреждений: ${result.specification.warnings.length}`;
    } catch (error) {
      previewTitle.textContent = 'Не удалось загрузить данные'; previewState.textContent = error.message || 'Ошибка связи';
    }
  }
  byId('previewButton').addEventListener('click', previewDeal);
  loadSpecificationsButton.addEventListener('click', loadSpecifications);
  dealId.addEventListener('change', () => {
    while (specificationSelect.options.length > 1) specificationSelect.remove(1);
    specificationSelect.selectedIndex = 0;
  });
  dealId.addEventListener('keydown', (event) => { if (event.key === 'Enter') { event.preventDefault(); previewDeal(); } });
  templateSelect.addEventListener('change', () => { syncOutputFormats(); if (!previewPanel.classList.contains('hidden') && dealId.value) previewDeal(); });
  syncOutputFormats();

  byId('checkConnection').addEventListener('click', async () => {
    const button = byId('checkConnection'); const initial = button.textContent; button.textContent = 'Проверяем…'; button.disabled = true;
    try {
      const response = await fetch('/settings/check', {method: 'POST'});
      const result = await response.json();
      button.textContent = result.message;
      setTimeout(() => { button.textContent = initial; button.disabled = false; }, 4200);
    } catch (_error) {
      button.textContent = 'Ошибка проверки'; setTimeout(() => { button.textContent = initial; button.disabled = false; }, 2500);
    }
  });
})();

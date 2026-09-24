/* Brightspace Quiz Converter - dependency-free browser implementation. */
const $ = (id) => document.getElementById(id);
const state = { file: null, quiz: null, assets: [] };
const MAX_FILE_SIZE = 10 * 1024 * 1024;

const fileInput = $('fileInput');
const dropZone = $('dropZone');
fileInput.addEventListener('change', () => fileInput.files.length && loadFiles(fileInput.files));
['dragenter', 'dragover'].forEach(event => dropZone.addEventListener(event, e => { e.preventDefault(); dropZone.classList.add('dragging'); }));
['dragleave', 'drop'].forEach(event => dropZone.addEventListener(event, e => { e.preventDefault(); dropZone.classList.remove('dragging'); }));
dropZone.addEventListener('drop', e => e.dataTransfer.files.length && loadFiles(e.dataTransfer.files));
$('clearFile').addEventListener('click', clearFile);
$('downloadButton').addEventListener('click', downloadPackage);
$('quizTitle').addEventListener('input', () => state.quiz && renderPreview(state.quiz));
$('quizDescription').addEventListener('input', () => state.quiz && renderPreview(state.quiz));

async function loadFile(file) { return loadFiles([file]); }
async function loadFiles(fileList) {
  hideError();
  const files = Array.from(fileList), file = files.find(candidate => /\.(txt|docx|csv)$/i.test(candidate.name));
  if (!file) return showError('Please choose a .docx or .txt quiz file.');
  if (files.some(candidate => candidate.size > MAX_FILE_SIZE)) return showError('One of those files is larger than the 10 MB limit.');
  try {
    const assetFiles = files.filter(candidate => candidate !== file && /^image\/(gif|jpe?g|png|svg\+xml)$/i.test(candidate.type || 'image/' + (candidate.name.split('.').pop() || '')));
    state.file = file; state.assets = await Promise.all(assetFiles.map(async asset => ({ name: asset.name, data: bytesToBase64(new Uint8Array(await asset.arrayBuffer())) })));
    $('fileName').textContent = file.name;
    $('fileSize').textContent = `${formatBytes(file.size)}${state.assets.length ? ` + ${state.assets.length} image${state.assets.length === 1 ? '' : 's'}` : ''}`;
    $('fileType').textContent = file.name.toLowerCase().endsWith('.docx') ? 'DOCX' : file.name.toLowerCase().endsWith('.csv') ? 'CSV' : 'TXT';
    $('fileRow').classList.remove('hidden');
    $('filePrompt').textContent = 'File loaded — click here to replace it';
    const text = file.name.toLowerCase().endsWith('.docx') ? await readDocx(file) : await file.text();
    const quiz = file.name.toLowerCase().endsWith('.csv') ? parseCsvQuiz(text) : parseQuiz(text);
    if (!quiz.questions.length) throw new Error('No questions were found. Check the recommended format below and try again.');
    quiz.assets = state.assets;
    state.quiz = quiz;
    if ($('quizTitle').value === 'Imported Quiz' && quiz.title) $('quizTitle').value = quiz.title;
    if (!$('quizDescription').value && quiz.description) $('quizDescription').value = quiz.description;
    renderPreview(quiz);
  } catch (error) {
    clearFile(false);
    showError(error.message || 'The file could not be read.');
  }
}

function clearFile(resetInput = true) {
  state.file = null; state.quiz = null; state.assets = [];
  if (resetInput) fileInput.value = '';
  $('fileRow').classList.add('hidden'); $('previewCard').classList.add('hidden');
  $('filePrompt').textContent = 'Drop a .docx or .txt file here'; hideError();
}

function showError(message) { $('fileError').textContent = message; $('fileError').classList.remove('hidden'); }
function hideError() { $('fileError').classList.add('hidden'); }
function formatBytes(bytes) { return bytes < 1024 ? `${bytes} B` : `${(bytes / 1024).toFixed(1)} KB`; }

// Extract readable paragraphs and embedded images from a DOCX (which is a ZIP file).
async function readDocx(file) {
  const entries = await readZip(new Uint8Array(await file.arrayBuffer()));
  const documentXml = decodeUtf8(entries['word/document.xml']);
  if (!documentXml) throw new Error('This Word document does not contain a readable document.xml file.');
  const relsXml = decodeUtf8(entries['word/_rels/document.xml.rels'] || new Uint8Array());
  const relationMap = {};
  if (relsXml) for (const rel of xmlDoc(relsXml).getElementsByTagName('Relationship')) {
    relationMap[rel.getAttribute('Id')] = (rel.getAttribute('Target') || '').replace(/^\/?/, 'word/');
  }
  const doc = xmlDoc(documentXml), body = doc.getElementsByTagName('body')[0];
  const lines = [], images = [];
  if (!body) return '';
  for (const node of body.children) {
    if (node.localName === 'tbl') {
      for (const row of Array.from(node.children).filter(n => n.localName === 'tr')) {
        const cells = Array.from(row.children).filter(n => n.localName === 'tc').map(cell => paragraphText(cell, relationMap, entries, images).trim());
        if (cells.some(Boolean)) lines.push(cells.join('\t'));
      }
    } else if (node.localName === 'p') {
      const value = paragraphText(node, relationMap, entries, images).trim();
      if (value || images.length) lines.push(value);
    }
  }
  // Image tokens allow the normal text parser to attach a DOCX image to the next question.
  return lines.join('\n') + (images.length ? `\n\n${JSON.stringify({ __quizImages: images.map(image => ({ name: image.name, data: bytesToBase64(image.data) })) })}` : '');
}

function paragraphText(node, relationMap, entries, images) {
  let text = '';
  for (const child of node.getElementsByTagName('*')) {
    if (child.localName === 't') text += child.textContent;
    else if (child.localName === 'tab') text += '\t';
    else if (child.localName === 'br' || child.localName === 'cr') text += '\n';
    else if (child.localName === 'blip') {
      const target = relationMap[child.getAttribute('r:embed')] || relationMap[child.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed')];
      if (target && entries[target]) { const name = `image-${images.length + 1}${extensionFor(target)}`; images.push({ name, data: entries[target] }); text += `\n[Image: ${name}]`; }
    }
  }
  return text;
}
function extensionFor(path) { const match = path.match(/\.[a-z0-9]+$/i); return match ? match[0].toLowerCase() : '.png'; }
function bytesToBase64(bytes) { let binary = ''; const chunk = 0x8000; for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk)); return btoa(binary); }

// Parser for the Respondus Standard Format plus the simpler format used by the original prototype.
function parseQuiz(source) {
  let images = [];
  const imageMarker = source.match(/\{"__quizImages":([\s\S]+)\}\s*$/);
  if (imageMarker) { try { images = JSON.parse(`[{"__quizImages":${imageMarker[1]}}]`)[0].__quizImages || []; } catch {} source = source.slice(0, imageMarker.index); }
  const rawLines = source.replace(/\r/g, '').split('\n').map(line => line.replace(/\u00a0/g, ' ').trim());
  const answerIndex = rawLines.findIndex(line => /^answers\s*:\s*$/i.test(line));
  const answerMap = answerIndex >= 0 ? parseAnswerList(rawLines.slice(answerIndex + 1)) : {};
  const lines = (answerIndex >= 0 ? rawLines.slice(0, answerIndex) : rawLines);
  const firstContent = lines.findIndex(Boolean);
  let title = '', description = '';
  const nextContent = firstContent >= 0 ? lines.slice(firstContent + 1).find(Boolean) || '' : '';
  if (firstContent >= 0 && /^(title|quiz title)\s*:/i.test(lines[firstContent]) && !/^type\s*:/i.test(nextContent)) { title = lines[firstContent].replace(/^[^:]+:\s*/i, ''); lines[firstContent] = ''; }
  const descriptionIndex = lines.findIndex(line => /^(description|instructions?)\s*:/i.test(line));
  if (descriptionIndex >= 0 && !lines.slice(0, descriptionIndex).some(isQuestionStart)) { description = lines[descriptionIndex].replace(/^[^:]+:\s*/i, ''); lines[descriptionIndex] = ''; }
  const questionStarts = [];
  let activeStructured = false, pendingStructured = false;
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index], typeLine = line.match(/^type\s*:\s*(.+)$/i);
    if (typeLine) {
      if (/^(mt|e|es|f|fb|mr|ma|tf|mc)$/i.test(typeLine[1].trim())) {
        activeStructured = false; pendingStructured = true;
      } else if (/^(matching|order|ordering|arrange)/i.test(typeLine[1])) {
        if (questionStarts.length && index > questionStarts[questionStarts.length - 1]) activeStructured = true;
        else pendingStructured = true;
      } else { activeStructured = false; pendingStructured = false; }
    }
    if (!isQuestionStart(line)) continue;
    const numberedRow = /^\d+\s*[.)\-:]/.test(line);
    if (pendingStructured) { questionStarts.push(index); activeStructured = true; pendingStructured = false; }
    else if (activeStructured && numberedRow && lines[index - 1] !== '') continue;
    else { questionStarts.push(index); activeStructured = /\b(match|matching|order|ordering|arrange)\b/i.test(cleanQuestionLine(line)); }
  }
  const blockStarts = questionStarts.map((questionStart, index) => {
    if (index === 0) return 0;
    let start = questionStart;
    while (start > questionStarts[index - 1] && (!lines[start - 1] || /^(type|title|points)\s*:/i.test(lines[start - 1]))) start--;
    return start;
  });
  const questions = [];
  for (let i = 0; i < questionStarts.length; i++) {
    const block = lines.slice(blockStarts[i], blockStarts[i + 1] ?? lines.length);
    const question = parseQuestionBlock(block, questions.length + 1, images);
    if (!question) continue;
    const listedAnswers = answerMap[question.number];
    if (listedAnswers?.length) {
      question.answerText = question.type === 'long-answer' ? listedAnswers.join('\n') : listedAnswers.join(', ');
      question.answers = question.type === 'fill-blank' ? listedAnswers : parseAnswers(question.answerText, question.options, question.type);
    }
    question.warning = answerWarning(question);
    questions.push(question);
  }
  let currentPoints = 1;
  for (const question of questions) { if (question.points != null) currentPoints = question.points; question.points = currentPoints; }
  if (!title && questions.length === 0 && firstContent >= 0) title = lines[firstContent];
  return { title, description, questions, images, assets: [] };
}

function parseAnswerList(lines) {
  const answers = {}, entries = lines.filter(Boolean); let current = null;
  for (const line of entries) {
    const match = line.match(/^(\d+)\s*[.)]\s*(.*)$/);
    if (match) { current = Number(match[1]); (answers[current] ||= []).push(match[2].trim()); }
    else if (current != null) answers[current][answers[current].length - 1] += ` ${line}`;
  }
  return answers;
}

function parseCsvQuiz(source) {
  const rows = parseCsvRows(source).filter(row => row.some(cell => cell.trim()));
  if (rows.length && /^type$/i.test(rows[0][0]?.trim())) rows.shift();
  const questions = rows.map((row, index) => {
    const type = normalizeType(row[0] || 'MC'), options = row.slice(5, 15).map((text, choiceIndex) => ({ label: String.fromCharCode(65 + choiceIndex), text: text.trim(), feedback: (row[19 + choiceIndex] || '').trim() })).filter(option => option.text);
    const answerText = (row[4] || '').trim(), prompt = (row[3] || '').trim();
    const answers = type === 'fill-blank' ? options.map(option => option.text) : parseAnswers(answerText, options, type);
    const feedback = [row[15], row[16], row[17]].filter(Boolean).join(' ').trim();
    return { number: index + 1, title: (row[1] || '').trim(), points: Number(row[2]) || 1, prompt, type, options, pairs: [], ordering: [], answers, answerText, feedback, imageRefs: parseImageRefs(prompt), warning: !answerText && !['fill-blank', 'long-answer'].includes(type) ? 'No answer found' : '' };
  });
  return { title: '', description: '', questions, images: [], assets: [] };
}
function parseCsvRows(source) {
  const rows = [], row = []; let field = '', quoted = false;
  for (let i = 0; i < source.length; i++) { const char = source[i], next = source[i + 1]; if (char === '"' && quoted && next === '"') { field += '"'; i++; } else if (char === '"') quoted = !quoted; else if (char === ',' && !quoted) { row.push(field); field = ''; } else if ((char === '\n' || char === '\r') && !quoted) { if (char === '\r' && next === '\n') i++; row.push(field); rows.push(row.splice(0)); field = ''; } else field += char; }
  if (field || row.length) { row.push(field); rows.push(row); } return rows;
}

function isQuestionStart(line) { return /^(?:question\s*)?\d+\s*[.)]\s*\S/i.test(line) || /^(?:q\s*\d+|question\s*\d+)\s*:/i.test(line); }
function questionNumber(line, fallback) { const match = line.match(/^(?:question\s*)?(\d+)\s*[.):\-]/i) || line.match(/^q\s*(\d+)\s*:/i); return match ? Number(match[1]) : fallback; }
function cleanQuestionLine(line) { return line.replace(/^(?:question\s*)?\d+\s*[.):\-]\s*/i, '').replace(/^q\s*\d+\s*:\s*/i, '').trim(); }
function parseQuestionBlock(block, fallbackNumber, images) {
  const questionIndex = block.findIndex(isQuestionStart); if (questionIndex < 0) return null;
  const questionLine = block[questionIndex];
  let prompt = cleanQuestionLine(questionLine), type = '', title = '', points = null, answerText = '', feedback = '', options = [], pairs = [], ordering = [], starAnswers = [], essayAnswer = '';
  let currentOption = null, mode = 'prompt', feedbackTarget = null;
  const beforeQuestion = block.slice(0, questionIndex);
  for (const line of beforeQuestion) {
    const meta = parseMetadata(line);
    if (meta.type) type = normalizeType(meta.type);
    if (meta.title) title = meta.title;
    if (meta.points != null) points = meta.points;
  }
  for (let i = questionIndex + 1; i < block.length; i++) {
    const line = block[i]; if (!line) continue;
    let match;
    const meta = parseMetadata(line);
    if (meta.type) { type = normalizeType(meta.type); mode = 'prompt'; continue; }
    if (meta.title) { title = meta.title; continue; }
    if (meta.points != null) { points = meta.points; continue; }
    if ((match = line.match(/^(?:answer|correct answer|correct)\s*:\s*(.+)$/i))) { answerText = match[1].trim(); mode = 'answer'; feedbackTarget = null; continue; }
    if ((match = line.match(/^~\s+(.+)$/))) { feedback += `${feedback ? ' ' : ''}${match[1].trim()}`; feedbackTarget = 'general'; continue; }
    if ((match = line.match(/^@\s+(.+)$/))) {
      if (currentOption) currentOption.feedback = `${currentOption.feedback ? ' ' : ''}${match[1].trim()}`;
      else feedback += `${feedback ? ' ' : ''}${match[1].trim()}`;
      feedbackTarget = currentOption ? 'option' : 'general'; continue;
    }
    if (type === 'long-answer' && (match = line.match(/^a\s*[.)]\s*(.+)$/i))) { essayAnswer = match[1].trim(); mode = 'answer'; continue; }
    if ((match = line.match(/^(\*)?\s*([A-Z])\s*[.)\-:]\s*(.+)$/i))) {
      if (type === 'matching' || /\b(match|matching)\b/i.test(prompt)) { const pair = match[3].match(/^(.+?)\s*=\s*(.+)$/); if (pair) pairs.push({ left: pair[1].trim(), right: pair[2].trim() }); else pairs.push({ left: match[3].trim(), right: '' }); continue; }
      const option = { label: match[2].toUpperCase(), text: match[3].trim(), feedback: '' }; options.push(option); currentOption = option;
      if (match[1]) starAnswers.push(options.length - 1);
      mode = 'options'; feedbackTarget = null; continue;
    }
    if ((match = line.match(/^(\d+)\s*[.)\-:]\s*(.+)$/))) {
      if (type === 'matching' || /\b(match|matching)\b/i.test(prompt)) {
        const pair = match[2].match(/^(.+?)\s*(?:=>|->|=)\s*(.+)$/); pairs.push(pair ? { left: pair[1].trim(), right: pair[2].trim() } : { left: match[2], right: '' });
      } else if (type === 'ordering' || /\b(order|ordering|arrange)\b/i.test(prompt)) ordering.push(match[2]);
      else if (mode === 'options') options.push({ label: match[1], text: match[2].trim(), feedback: '' });
      continue;
    }
    if (line.includes('\t') && (type === 'matching' || /\b(match|matching)\b/i.test(prompt))) { const [left, right] = line.split('\t'); pairs.push({ left: left.trim(), right: (right || '').trim() }); continue; }
    if ((match = line.match(/^(.+?)\s*(?:=>|->|=)\s*(.+)$/)) && (type === 'matching' || /\b(match|matching)\b/i.test(prompt))) { pairs.push({ left: match[1].trim(), right: match[2].trim() }); continue; }
    if (feedbackTarget === 'option' && currentOption) currentOption.feedback += ` ${line}`;
    else if (feedbackTarget === 'general') feedback += ` ${line}`;
    else if (type === 'long-answer' && mode === 'answer') essayAnswer += ` ${line}`;
    else if (mode === 'options' && currentOption) currentOption.text += ` ${line}`;
    else prompt += ` ${line}`;
  }
  if (!type) type = inferType(prompt, options, pairs, ordering, answerText);
  if (type === 'true-false' && !options.length) options = [{ label: 'A', text: 'True', feedback: '' }, { label: 'B', text: 'False', feedback: '' }];
  if (type === 'ordering' && !ordering.length) ordering = options.map(o => o.text);
  if (type === 'matching' && !pairs.length && options.length) pairs = options.map(o => ({ left: o.text, right: o.text }));
  if (type === 'fill-blank' && options.length && !answerText) answerText = options.map(option => option.text).join(' | ');
  if (type === 'long-answer' && essayAnswer && !answerText) answerText = essayAnswer;
  const answers = starAnswers.length ? starAnswers : parseAnswers(answerText, options, type);
  const imageRefs = [];
  imageRefs.push(...parseImageRefs(prompt));
  prompt = prompt.replace(/\[Image:\s*[^\]]+\]/gi, '').replace(/\[\s*img:[^\]]+\]/gi, '').trim();
  const warning = !answerText && !starAnswers.length && ['multiple-choice', 'true-false', 'multi-select', 'fill-blank'].includes(type) ? 'No answer found' : '';
  return { number: questionNumber(questionLine, fallbackNumber), title, points, prompt, type, options, pairs, ordering, answers, answerText, feedback, imageRefs, warning };
}
function parseImageRefs(text) { const refs = []; for (const match of text.matchAll(/\[Image:\s*([^\]]+)\]/gi)) refs.push({ name: match[1].trim(), alt: 'Quiz image' }); for (const match of text.matchAll(/\[\s*img:\s*["“]([^"”]+)["”](?:\s+["“]([^"”]+)["”])?\s*\]/gi)) refs.push({ name: match[1].trim(), alt: match[2]?.trim() || 'Quiz image' }); return refs; }
function answerWarning(question) { return !question.answerText && !question.answers.length && ['multiple-choice', 'true-false', 'multi-select', 'fill-blank'].includes(question.type) ? 'No answer found' : ''; }
function parseMetadata(line) { const type = line.match(/^type\s*:\s*(.+)$/i); const title = line.match(/^title\s*:\s*(.+)$/i); const points = line.match(/^points\s*:\s*([\d.]+)/i); return { type: type?.[1], title: title?.[1]?.trim(), points: points ? Number(points[1]) : null }; }
function normalizeType(value) { const v = value.toLowerCase().trim(); if (/^(mr|ma)$|multi.?select|multiple answers|multiple response/.test(v)) return 'multi-select'; if (/^(tf)$|true.?false/.test(v)) return 'true-false'; if (/^(f|fb)$|fill|blank|fib/.test(v)) return 'fill-blank'; if (/^(mt)$|match/.test(v)) return 'matching'; if (/order|arrange/.test(v)) return 'ordering'; if (/^(e|es)$|long|essay|written|paragraph/.test(v)) return 'long-answer'; return 'multiple-choice'; }
function inferType(prompt, options, pairs, ordering, answer) { const optionWords = options.map(o => o.text.toLowerCase()); if (optionWords.length >= 2 && /^(true|t)$/.test(optionWords[0]) && /^(false|f)$/.test(optionWords[1])) return 'true-false'; if (/\btrue\s*\/\s*false\b|\btrue or false\b/i.test(prompt)) return 'true-false'; if (pairs.length || /\bmatching\b/i.test(prompt)) return 'matching'; if (ordering.length || /\b(?:ordering|arrange in order)\b/i.test(prompt)) return 'ordering'; if (/\b(?:fill in the blank|fill-in-the-blank)\b/i.test(prompt)) return 'fill-blank'; if (/\b(?:long answer|essay|short answer)\b/i.test(prompt)) return 'long-answer'; if (answer.includes(',') || answer.includes(' and ') || /^[A-Z](?:[\s,]+[A-Z])+$/i.test(answer.trim())) return 'multi-select'; return 'multiple-choice'; }
function parseAnswers(text, options, type = '') { if (!text) return []; let values = text.replace(/[“”]/g, '').trim().split(/\s*(?:,|;|\band\b)\s*/i).filter(Boolean); if (values.length === 1 && options.length && /^(?:[A-J]|\d+)(?:\s+(?:[A-J]|\d+))+$/i.test(values[0])) values = values[0].split(/\s+/); return values.map(value => { const normalized = value.trim(); if (type === 'true-false' || (options.length === 2 && /^(true|false|t|f|a|b)$/i.test(normalized))) { if (/^(true|t|a)$/i.test(normalized)) return 0; if (/^(false|f|b)$/i.test(normalized)) return 1; } const match = normalized.match(/^([A-J]|\d+)\b/i); if (match && options.length) { const byLabel = options.findIndex(o => o.label.toUpperCase() === match[1].toUpperCase()); return byLabel >= 0 ? byLabel : Math.max(0, Number(match[1]) - 1); } const byText = options.findIndex(o => o.text.trim().toLowerCase() === normalized.toLowerCase()); return byText >= 0 ? byText : normalized; }).filter(v => v !== ''); }

function renderPreview(quiz) {
  $('previewCard').classList.remove('hidden');
  $('summaryText').textContent = `${quiz.questions.length} question${quiz.questions.length === 1 ? '' : 's'} detected · review before downloading`;
  const availableImages = new Set([...(quiz.images || []), ...(quiz.assets || [])].map(image => image.name));
  const warnings = quiz.questions.flatMap(q => {
    const messages = q.warning ? [`Question ${q.number}: ${q.warning}.`] : [];
    const missing = (q.imageRefs || []).filter(ref => !availableImages.has(ref.name)).map(ref => ref.name);
    if (missing.length) messages.push(`Question ${q.number}: image file${missing.length === 1 ? '' : 's'} not included (${missing.join(', ')}).`);
    return messages;
  });
  $('warningBox').innerHTML = warnings.length ? `<strong>Review recommended</strong><ul>${warnings.map(escapeHtml).map(w => `<li>${w}</li>`).join('')}</ul>` : '';
  $('warningBox').classList.toggle('hidden', !warnings.length);
  $('questionList').innerHTML = quiz.questions.map(q => `<div class="question-item"><span class="question-number">${String(q.number).padStart(2, '0')}</span><span class="question-text">${escapeHtml(q.prompt || '(blank question)')}</span><span class="question-meta">${escapeHtml(labelForType(q.type))}</span></div>`).join('');
}
function labelForType(type) { return ({ 'multiple-choice': 'Multiple choice', 'true-false': 'True / False', 'fill-blank': 'Fill in blank', 'multi-select': 'Multi-select', matching: 'Matching', ordering: 'Ordering', 'long-answer': 'Long answer' })[type] || type; }

async function downloadPackage() {
  if (!state.quiz) return;
  const quiz = { ...state.quiz, title: $('quizTitle').value.trim() || 'Imported Quiz', description: $('quizDescription').value.trim() };
  const files = buildBrightspaceFiles(quiz);
  const blob = new Blob([createZip(files)], { type: 'application/zip' });
  const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = `${safeName(quiz.title)}-brightspace.zip`; link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function buildBrightspaceFiles(quiz) {
  const quizId = `quiz_d2l_${Date.now()}`;
  const resourceId = `res_quiz_${Date.now()}`;
  const images = [...(quiz.images || []), ...(quiz.assets || [])];
  const imageNames = new Set(images.map(i => i.name));
  const xml = buildQuizXml(quiz, resourceId, imageNames);
  const manifest = `<?xml version="1.0" encoding="UTF-8"?>\n<manifest identifier="D2L_${xmlId()}" xmlns:d2l_2p0="http://desire2learn.com/xsd/d2lcp_v2p0" xmlns:imsmd="http://www.imsglobal.org/xsd/imsmd_rootv1p2p1" xmlns="http://www.imsglobal.org/xsd/imscp_v1p1"><metadata><imsmd:lom><imsmd:general><imsmd:title><imsmd:langstring xml:lang="en-us">${xmlEscape(quiz.title)}</imsmd:langstring></imsmd:title><imsmd:language>en-us</imsmd:language></imsmd:general></imsmd:lom></metadata><resources><resource identifier="${resourceId}" type="webcontent" d2l_2p0:material_type="d2lquiz" d2l_2p0:link_target="" href="${quizId}.xml" title="${xmlEscape(quiz.title)}" /></resources></manifest>`;
  const files = [{ name: 'imsmanifest.xml', data: utf8(manifest) }, { name: `${quizId}.xml`, data: utf8(xml) }];
  for (const image of images) if (imageNames.has(image.name) && !files.some(file => file.name === image.name)) files.push({ name: image.name, data: Uint8Array.from(atob(image.data), c => c.charCodeAt(0)) });
  return files;
}

function buildQuizXml(quiz, resourceId, imageNames) {
  const ns = 'http://desire2learn.com/xsd/d2lcp_v2p0';
  const items = quiz.questions.map((q, index) => buildItem(q, index + 1, imageNames)).join('');
  return `<?xml version="1.0" encoding="UTF-8"?><questestinterop xmlns:d2l_2p0="${ns}"><assessment d2l_2p0:id="1" title="${xmlEscape(quiz.title)}" ident="${resourceId}" d2l_2p0:resource_code="${xmlId()}"><rubric><flow_mat><material><mattext d2l_2p0:isdisplayed="yes" texttype="text/html">${htmlMattext(quiz.description)}</mattext></material></flow_mat></rubric><assessmentcontrol hide_question_pointsswitch="no" hintswitch="no" solutionswitch="no" feedbackswitch="no" /><presentation_material><flow_mat><material label="page header"><mattext d2l_2p0:isdisplayed="yes" texttype="text/html" /></material><material label="page footer"><mattext d2l_2p0:isdisplayed="yes" texttype="text/html" /></material></flow_mat></presentation_material><assess_procextension><d2l_2p0:intro_message d2l_2p0:isdisplayed="no" texttype="text/plain" /><d2l_2p0:disable_right_click>no</d2l_2p0:disable_right_click><d2l_2p0:disable_pager_access>no</d2l_2p0:disable_pager_access><is_active>no</is_active><d2l_2p0:time_limit>0</d2l_2p0:time_limit><d2l_2p0:show_clock>no</d2l_2p0:show_clock><d2l_2p0:attempts_allowed>1</d2l_2p0:attempts_allowed><d2l_2p0:mark_calculation_type>1</d2l_2p0:mark_calculation_type><d2l_2p0:is_forward_only>no</d2l_2p0:is_forward_only><d2l_2p0:paging_type_id>0</d2l_2p0:paging_type_id></assess_procextension><assessfeedback><rubric><flow_mat><material><mattext texttype="no" /></material></flow_mat></rubric><d2l_2p0:duration>0</d2l_2p0:duration><d2l_2p0:response_display_type_id>1</d2l_2p0:response_display_type_id><d2l_2p0:show_correct_answers>no</d2l_2p0:show_correct_answers><d2l_2p0:submission_restrictip>no</d2l_2p0:submission_restrictip><d2l_2p0:show_class_average>no</d2l_2p0:show_class_average><d2l_2p0:show_score_distribution>no</d2l_2p0:show_score_distribution></assessfeedback><section ident="CONTAINER_SECTION">${items}</section></assessment></questestinterop>`;
}

function buildItem(q, index, imageNames) {
  const id = `QUES_${xmlId()}_${index}`, lid = `${id}_LID`, answerIds = q.options.map((_, i) => `${id}_A${i + 1}`);
  const title = q.title ? ` title="${xmlEscape(q.title)}"` : '';
  const prompt = addImageHtml(q.prompt, q.imageRefs, imageNames);
  const metadata = `<itemmetadata><qtimetadata><qti_metadatafield><fieldlabel>qmd_computerscored</fieldlabel><fieldentry>${q.type === 'long-answer' ? 'no' : 'yes'}</fieldentry></qti_metadatafield><qti_metadatafield><fieldlabel>qmd_questiontype</fieldlabel><fieldentry>${xmlEscape(labelForType(q.type))}</fieldentry></qti_metadatafield><qti_metadatafield><fieldlabel>qmd_weighting</fieldlabel><fieldentry>${Number(q.points ?? 1).toFixed(9)}</fieldentry></qti_metadatafield></qtimetadata></itemmetadata><itemproc_extension><d2l_2p0:difficulty>1</d2l_2p0:difficulty><d2l_2p0:isbonus>no</d2l_2p0:isbonus><d2l_2p0:ismandatory>no</d2l_2p0:ismandatory></itemproc_extension>`;
  if (q.type === 'long-answer') return `<item ident="OBJ_${xmlId()}" label="${id}"${title} d2l_2p0:page="1">${metadata}<presentation><flow><material><mattext texttype="text/html">${htmlMattext(prompt)}</mattext></material><response_extension><d2l_2p0:has_signed_comments>no</d2l_2p0:has_signed_comments><d2l_2p0:has_htmleditor>yes</d2l_2p0:has_htmleditor><d2l_2p0:has_fileupload>no</d2l_2p0:has_fileupload></response_extension><response_str ident="${id}_STR" rcardinality="Multiple"><render_fib rows="15" columns="100" prompt="Box" fibtype="String"><response_label ident="${id}_LA" /></render_fib></response_str></flow></presentation></item>`;
  if (q.type === 'fill-blank') return buildFillBlank(q, id, metadata, title, prompt);
  if (q.type === 'matching') return buildMatching(q, id, metadata, title, prompt);
  if (q.type === 'ordering') return buildOrdering(q, id, metadata, title, prompt);
  const multi = q.type === 'multi-select', opts = (q.options.length ? q.options : [{ label: 'A', text: 'True' }, { label: 'B', text: 'False' }]);
  const labels = opts.map((o, i) => `<flow_label class="Block"><response_label ident="${answerIds[i]}"><flow_mat><material><mattext texttype="text/html">${htmlMattext(o.text)}</mattext></material></flow_mat></response_label></flow_label>`).join('');
  const conditions = opts.map((o, i) => `<respcondition title="Response Condition ${i + 1}"><conditionvar><varequal respident="${lid}">${answerIds[i]}</varequal></conditionvar><setvar action="Set">${q.answers.includes(i) ? '100.000000000' : '0.000000000'}</setvar>${o.feedback ? `<displayfeedback feedbacktype="Response" linkrefid="${id}_IF${i + 1}" />` : ''}</respcondition>`).join('');
  const optionFeedback = opts.map((o, i) => o.feedback ? `<itemfeedback ident="${id}_IF${i + 1}"><material><mattext texttype="text/html">${htmlMattext(o.feedback)}</mattext></material></itemfeedback>` : '').join('');
  return `<item ident="OBJ_${xmlId()}" label="${id}"${title} d2l_2p0:page="1">${metadata}<presentation><flow><material><mattext texttype="text/html">${htmlMattext(prompt)}</mattext></material><response_extension><d2l_2p0:display_style>2</d2l_2p0:display_style><d2l_2p0:enumeration>6</d2l_2p0:enumeration><d2l_2p0:grading_type>0</d2l_2p0:grading_type></response_extension><response_lid ident="${lid}" rcardinality="${multi ? 'Multiple' : 'Single'}"><render_choice shuffle="no">${labels}</render_choice></response_lid></flow></presentation><resprocessing>${conditions}</resprocessing>${q.feedback ? `<itemfeedback ident="${id}_FEEDBACK"><material><mattext texttype="text/html">${htmlMattext(q.feedback)}</mattext></material></itemfeedback>` : ''}${optionFeedback}</item>`;
}

function buildFillBlank(q, id, metadata, title, prompt) { const answers = q.answers.length ? q.answers : [q.answerText || '']; return `<item ident="OBJ_${xmlId()}" label="${id}"${title} d2l_2p0:page="1">${metadata}<presentation><flow><material><mattext texttype="text/html">${htmlMattext(prompt)}</mattext></material><response_str ident="${id}_STR" rcardinality="Single"><render_fib rows="1" columns="30" prompt="Box" fibtype="String"><response_label ident="${id}_ANS" /></render_fib></response_str></flow></presentation><resprocessing><respcondition><conditionvar>${answers.map(a => `<varequal respident="${id}_ANS" case="no">${xmlEscape(String(a))}</varequal>`).join('')}</conditionvar><setvar action="Set">100.000000000</setvar></respcondition></resprocessing></item>`; }
function buildMatching(q, id, metadata, title, prompt) { const matches = q.pairs.length ? q.pairs : [{ left: 'Match', right: 'Answer' }], choices = [...new Set(matches.map(p => p.right))]; const groups = matches.map((p, i) => `<response_grp respident="${id}_C${i + 1}" rcardinality="Single"><material><mattext texttype="text/html">${htmlMattext(p.left)}</mattext></material><render_choice shuffle="yes">${choices.map((c, j) => `<flow_label class="Block"><response_label ident="${id}_M${j + 1}"><flow_mat><material><mattext texttype="text/html">${htmlMattext(c)}</mattext></material></flow_mat></response_label></flow_label>`).join('')}</render_choice></response_grp>`).join(''); const conditions = matches.map((p, i) => `<respcondition><conditionvar><varequal respident="${id}_C${i + 1}">${id}_M${choices.indexOf(p.right) + 1}</varequal></conditionvar><setvar varname="D2L_Correct" action="Add">1</setvar></respcondition>`).join(''); return `<item ident="OBJ_${xmlId()}" label="${id}"${title} d2l_2p0:page="1">${metadata}<presentation><flow><material><mattext texttype="text/html">${htmlMattext(prompt)}</mattext></material>${groups}</flow></presentation><resprocessing><outcomes><decvar vartype="Integer" defaultval="0" varname="D2L_Correct" /></outcomes>${conditions}</resprocessing></item>`; }
function buildOrdering(q, id, metadata, title, prompt) { const values = q.ordering.length ? q.ordering : ['First', 'Second']; const labels = values.map((value, i) => `<flow_label class="Block"><response_label ident="${id}_O${i + 1}"><flow_mat><material><mattext texttype="text/html">${htmlMattext(value)}</mattext></material></flow_mat></response_label></flow_label>`).join(''); const conditions = values.map((_, i) => `<respcondition><conditionvar><varequal respident="${id}_O${i + 1}">${i + 1}</varequal></conditionvar><setvar varname="D2L_Correct" action="Add">1</setvar></respcondition>`).join(''); return `<item ident="OBJ_${xmlId()}" label="${id}"${title} d2l_2p0:page="1">${metadata}<presentation><flow><material><mattext texttype="text/html">${htmlMattext(prompt)}</mattext></material><response_grp respident="${id}_O" rcardinality="Ordered"><render_choice shuffle="yes">${labels}</render_choice></response_grp></flow></presentation><resprocessing><outcomes><decvar vartype="Integer" defaultval="0" varname="D2L_Correct" /></outcomes>${conditions}</resprocessing></item>`; }

function addImageHtml(text, refs, imageNames) { const images = (refs || []).map(ref => `<img src="${xmlEscape(ref.name)}" alt="${xmlEscape(ref.alt || 'Quiz image')}" style="max-width: 100%;"/>`).join(''); return images ? `${text} ${images}` : text; }
function htmlMattext(value) { const text = String(value || '').replace(/\[HTML\]/gi, '').replace(/\[\/HTML\]/gi, '').trim(); return xmlEscape(`<p>${text.replace(/\n+/g, '</p><p>')}</p>`); }
function xmlEscape(value) { return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;'); }
function escapeHtml(value) { return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;'); }
function safeName(value) { return value.replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase() || 'quiz'; }
function xmlId() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }

// Small ZIP reader/writer. DOCX commonly uses deflate; generated packages use no compression.
async function readZip(bytes) { const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength), entries = {}; let eocd = -1; for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65558); i--) if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; } if (eocd < 0) throw new Error('The Word file is not a valid ZIP package.'); const count = view.getUint16(eocd + 10, true), offset = view.getUint32(eocd + 16, true); let p = offset; for (let n = 0; n < count; n++) { if (view.getUint32(p, true) !== 0x02014b50) break; const method = view.getUint16(p + 10, true), compressed = view.getUint32(p + 20, true), nameLength = view.getUint16(p + 28, true), extraLength = view.getUint16(p + 30, true), commentLength = view.getUint16(p + 32, true), localOffset = view.getUint32(p + 42, true); const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLength)); const lp = localOffset, localNameLength = view.getUint16(lp + 26, true), localExtraLength = view.getUint16(lp + 28, true); const dataStart = lp + 30 + localNameLength + localExtraLength, packed = bytes.subarray(dataStart, dataStart + compressed); entries[name] = method === 0 ? packed : await inflateRaw(packed); p += 46 + nameLength + extraLength + commentLength; } return entries; }
async function inflateRaw(bytes) { if (!('DecompressionStream' in window)) throw new Error('This browser cannot decompress Word files. Please use a current version of Chrome, Edge, Firefox, or Safari.'); const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw')); return new Uint8Array(await new Response(stream).arrayBuffer()); }
function createZip(files) { const chunks = [], central = []; let offset = 0; for (const file of files) { const name = utf8(file.name), data = file.data instanceof Uint8Array ? file.data : utf8(file.data), crc = crc32(data); const local = concat([u32(0x04034b50), u16(20), u16(0x800), u16(0), u16(0), u16(0), u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), name, data]); chunks.push(local); central.push(concat([u32(0x02014b50), u16(20), u16(20), u16(0x800), u16(0), u16(0), u16(0), u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), name])); offset += local.length; } const directory = concat(central); return concat([...chunks, directory, u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length), u32(directory.length), u32(offset), u16(0)]); }
function utf8(value) { return new TextEncoder().encode(value); }
function u16(value) { const a = new Uint8Array(2); new DataView(a.buffer).setUint16(0, value, true); return a; }
function u32(value) { const a = new Uint8Array(4); new DataView(a.buffer).setUint32(0, value >>> 0, true); return a; }
function concat(arrays) { const result = new Uint8Array(arrays.reduce((n, a) => n + a.length, 0)); let at = 0; for (const a of arrays) { result.set(a, at); at += a.length; } return result; }
function crc32(bytes) { let crc = 0xffffffff; for (const byte of bytes) { crc ^= byte; for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1)); } return (crc ^ 0xffffffff) >>> 0; }
function decodeUtf8(bytes) { return bytes && bytes.length ? new TextDecoder().decode(bytes) : ''; }
function xmlDoc(text) { return new DOMParser().parseFromString(text, 'application/xml'); }

const SAMPLE = {
  title: 'The Attention Economy',
  slides: [
    { title: 'The Attention Economy', text: ['Attention is the scarce resource of the digital age.', 'Platforms compete to capture, measure, and resell human focus.'], notes: 'Use this as the opening thesis.' },
    { title: 'A simple exchange', text: ['Users receive free tools and social connection.', 'Advertisers receive a prediction about what users will do next.'], notes: '' },
    { title: 'What to ask', text: ['Who benefits from the design?', 'What behavior is being optimized?', 'Which voices are missing from the dataset?'], notes: 'Connect this to your media literacy reading.' }
  ], links: [], images: []
};

const state = { deck: null, view: 'guide' };
const $ = (selector) => document.querySelector(selector);

function extractId(value) {
  const match = value.match(/\/presentation\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

function textFromPageElement(element, links, images) {
  const output = [];
  const shape = element.shape;
  if (shape?.text?.textElements) {
    output.push(shape.text.textElements.map((item) => item.textRun?.content || '').join('').trim());
  }
  if (element.table) output.push('[Table with content]');
  if (element.image) images.push(element.image.contentUrl || element.image.sourceUrl || 'Image');
  const url = shape?.shapeProperties?.link?.url || shape?.text?.textElements?.find((item) => item.textRun?.style?.link?.url)?.textRun?.style?.link?.url;
  if (url) links.push(url);
  return output.filter(Boolean);
}

function parsePresentation(data) {
  const links = [], images = [], keyTerms = [];
  const slides = (data.slides || []).map((page, index) => {
    const text = [];
    (page.pageElements || []).forEach((element) => text.push(...textFromPageElement(element, links, images)));
    const noteElements = page.slideProperties?.notesPage?.pageElements || [];
    const notes = [];
    noteElements.forEach((element) => notes.push(...textFromPageElement(element, links, images)));
    const title = text[0] || `Slide ${index + 1}`;
    return { title, text, notes: notes.join(' ') };
  });
  return { title: data.title || 'Untitled presentation', slides, links, images, keyTerms };
}

async function fetchSingleDeck(url) {
  const id = extractId(url);
  const apiKey = $('#api-key').value.trim();
  const token = $('#access-token').value.trim();
  if (!id) throw new Error('That does not look like a Google Slides URL. Use the link from the presentation address bar.');
  if (!apiKey && !token) {
    const exportResponse = await fetch(`https://docs.google.com/presentation/d/${id}/export/pptx`);
    if (!exportResponse.ok) throw new Error('Google did not allow a public export. Set the deck to “Anyone with the link” or use the school-account .pptx upload.');
    const deck = await parsePptx(new File([await exportResponse.blob()], 'Google Slides export.pptx', { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
    deck.title = deck.slides[0]?.title || `Slides ${id.slice(0, 8)}`;
    return deck;
  }
  const endpoint = new URL(`https://slides.googleapis.com/v1/presentations/${id}`);
  if (apiKey) endpoint.searchParams.set('key', apiKey);
  const response = await fetch(endpoint, { headers: token ? { Authorization: `Bearer ${token}` } : {} });
  if (!response.ok) {
    const detail = await response.json().catch(() => ({}));
    throw new Error(detail.error?.message || `Google returned ${response.status}. Check the link and credentials.`);
  }
  return parsePresentation(await response.json());
}
function combineDecks(decks) {
  const validDecks = decks.filter(Boolean);
  const sourceNames = validDecks.map((deck, index) => { const titleSlide = deck.slides[0]?.title?.trim(); const title = titleSlide && !/^slide \d+$/i.test(titleSlide) ? titleSlide : deck.title; return validDecks.length > 1 ? `${title} · source ${index + 1}` : title; });
  return { title: validDecks.length === 1 ? sourceNames[0] : `${validDecks.length} source files`, sources: sourceNames, slides: validDecks.flatMap((deck, index) => deck.slides.map((slide) => ({ ...slide, source: sourceNames[index] }))), links: validDecks.flatMap((deck) => deck.links), images: validDecks.flatMap((deck) => deck.images), keyTerms: validDecks.flatMap((deck) => deck.keyTerms || []) };
}
async function fetchDeck() {
  const urls = [...document.querySelectorAll('.slides-url')].map((input) => input.value.trim()).filter(Boolean);
  if (!urls.length) throw new Error('Paste at least one Google Slides URL.');
  return combineDecks(await Promise.all(urls.map(fetchSingleDeck)));
}

async function parsePptx(file) {
  if (!window.JSZip) throw new Error('The PowerPoint parser could not load. Check your internet connection and reload the page.');
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const slideNames = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));
  const slides = [];
  for (const name of slideNames) {
    const xml = await zip.files[name].async('text');
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const paragraphs = [...document.getElementsByTagName('a:p')].map((paragraph) => [...paragraph.getElementsByTagName('a:t')].map((node) => node.textContent).join('').replace(/\s+/g, ' ').trim()).filter(Boolean);
    const text = paragraphs.length ? paragraphs : [...document.getElementsByTagName('a:t')].map((node) => node.textContent.trim()).filter(Boolean);
    const emphasized = [...document.getElementsByTagName('a:r')].filter((run) => { const style = run.getElementsByTagName('a:rPr')[0]; return style && (style.getAttribute('b') === '1' || style.getElementsByTagName('a:highlight').length); }).map((run) => run.getElementsByTagName('a:t')[0]?.textContent.trim()).filter(Boolean);
    slides.push({ title: text[0] || `Slide ${slides.length + 1}`, text, notes: '', emphasized });
  }
  if (!slides.length) throw new Error('No slides were found in that PowerPoint file.');
  const links = slides.flatMap((slide) => slide.text.filter((line) => /^https?:\/\//i.test(line)));
  return { title: file.name.replace(/\.pptx$/i, ''), slides, links, images: [], keyTerms: slides.flatMap((slide) => slide.emphasized) };
}

async function parseUploadedFile(file) {
  if (!file) return;
  if (file.name.toLowerCase().endsWith('.pptx')) return parsePptx(file);
  if (/\.(pdf|jpe?g|png|webp)$/i.test(file.name)) return parseScannedFile(file);
  const content = await file.text();
  if (file.name.toLowerCase().endsWith('.json')) return parsePresentation(JSON.parse(content));
  const slides = content.split(/\n\s*\n/).map((block, index) => { const text = block.split('\n').map((line) => line.trim()).filter(Boolean); return { title: text[0] || `Slide ${index + 1}`, text, notes: '' }; });
  return { title: file.name.replace(/\.[^.]+$/, '') || file.name, slides, links: [], images: [] };
}
async function parseUploadedFiles(files) {
  const decks = await Promise.all([...files].map(parseUploadedFile));
  return combineDecks(decks);
}

async function recognizeImage(source, label) {
  if (!window.Tesseract) throw new Error('The OCR engine could not load. Check your internet connection and reload the page.');
  const result = await Tesseract.recognize(source, 'eng', { logger: (message) => { if (message.status === 'recognizing text') $('#url-hint').textContent = `Reading ${label}: ${Math.round((message.progress || 0) * 100)}%`; } });
  return result.data.text.split('\n').map((line) => line.trim()).filter(Boolean);
}

async function parseScannedFile(file) {
  if (/\.pdf$/i.test(file.name)) {
    if (!window.pdfjsLib) throw new Error('The PDF parser could not load. Check your internet connection and reload the page.');
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.worker.min.mjs';
    const pdf = await window.pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    const slides = [];
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      const text = await recognizeImage(canvas, `page ${pageNumber} of ${pdf.numPages}`);
      slides.push({ title: text[0] || `Page ${pageNumber}`, text, notes: '' });
    }
    return { title: file.name.replace(/\.pdf$/i, '') || file.name, slides, links: [], images: [] };
  }
  const text = await recognizeImage(file, file.name);
  return { title: file.name.replace(/\.[^.]+$/, '') || file.name, slides: [{ title: text[0] || file.name, text, notes: '' }], links: [], images: [] };
}

function escapeHtml(value) { return String(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[char])); }
function setError(message) { $('#error-box').textContent = message; $('#error-box').classList.remove('is-hidden'); }
function clearError() { $('#error-box').classList.add('is-hidden'); }
function setLoading(loading) { const button = $('#mine-button'); button.disabled = loading; button.querySelector('span').textContent = loading ? 'Mining...' : 'Mine deck'; }

function renderStats(deck) {
  const textCount = deck.slides.reduce((count, slide) => count + slide.text.join(' ').split(/\s+/).filter(Boolean).length, 0);
  $('#stats').innerHTML = [['slides', deck.slides.length], ['words', textCount], ['links', deck.links.length], ['media', deck.images.length]].map(([label, value]) => `<div class="stat"><strong>${value}</strong><span>${label}</span></div>`).join('');
}
function contentLines(deck) { return deck.slides.flatMap((slide) => slide.text).map((line) => line.replace(/^[-•*]\s*/, '').replace(/\s+/g, ' ').trim()).filter(Boolean); }
function isPersonDescription(line) { return /\b(?:head of|outreach member|impact presenter|mech member|vp of|class of)\b/i.test(line) || /^(?:i|my|love|like|draw|dance|knit)\b/i.test(line) || /^[A-Z][a-z]+\s*[-–]\s*/.test(line); }
function isLowValuePoint(line) { return /[“”\"]/.test(line) || /^(?:hi|hello|hey|thanks|thank you|welcome|example|scenario|now|today|sit|feel free|you will|your job|please|what is|who are|meet the|masterclass|tips)\b/i.test(line) || /(?:^|:)\s*(?:example|masterclass|tips|be gp)\b/i.test(line) || /\b(?:attendance|fill it out|take a picture|screenshot|move to|split yourselves|grab a box|ask for help|use provided tools|click|right click|left click|scroll wheel|navigate to|if they ask you something you don’t know|if they ask you something you don't know|you finish helping them)\b/i.test(line); }
function looksLikePoint(line) { return line.length >= 28 && !line.endsWith('?') && !isPersonDescription(line) && !isLowValuePoint(line) && !/\b(?:important\s+first\s+value|steps\s+for\s+starting)\b/i.test(line) && !/^\(?\s*(?:by|about me|who are we|meet|what is)\b/i.test(line) && !/^\(?[A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3}\)?$/.test(line); }
function hasQuantitativeDetail(line) { return /\b\d+(?:[.,]\d+)?\s*(?:%|minutes?|hours?|days?|weeks?|months?|years?|in\.?|mm|cm|m|kg|lb|lbs|people|students?|groups?|parts?|steps?|channels?|ports?)?\b|\b(?:one|two|three|four|five|six|seven|eight|nine|ten|first|second|third)\b/i.test(line); }
function rankPoints(points) { return points.sort((left, right) => Number(hasQuantitativeDetail(right)) - Number(hasQuantitativeDetail(left))); }
function extractPeople(deck) {
  const lines = contentLines(deck);
  const people = [];
  lines.forEach((line, index) => {
    const match = line.match(/^(?:by\s+|about me:\s*)?([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,2})(?:\s*\(([^)]+)\)|\s*[-:]\s*(.+))?$/i);
    const hasContext = /^(?:by\s+|about me:\s*)/i.test(line) || /[-:()]/.test(line);
    if (!hasContext) return;
    if (!match) return;
    const blockedName = /^(?:meet|the|vp|head|library|what|who|steps|workshop|council|days|masterclass|tips|be)(?:\s|$)/i.test(match[1]);
    if (blockedName) return;
    const explicitPerson = /^(?:by\s+|about me:\s*)/i.test(line);
    const roleClue = /\b(?:head|member|presenter|vp|director|manager|lead|coordinator|captain|engineer|founder|chair)\b/i.test(match[2] || match[3] || '');
    if (!explicitPerson && !roleClue) return;
    const nearbyLines = lines.slice(index + 1, index + 20);
    const next = nearbyLines[0] || '';
    const nearbyRole = nearbyLines.find((nearbyLine) => /^(?:(?:vp|head|director|manager|lead|member|presenter)\b[^()]*)/i.test(nearbyLine))?.replace(/\s*\([^)]*\)/, '').trim() || '';
    const role = match[2] || match[3] || (next.match(/^\(([^)]+)\)$/)?.[1]) || nearbyRole;
    if (people.some((person) => person.name === match[1])) return;
    people.push({ name: match[1], role: role || 'Role not stated in the deck' });
  });
  return people.slice(0, 6);
}
function extractKeyTerms(deck, people) {
  const stopWords = new Set('about after again also because been being between booth can could first from have into interested look more most other over see seems should some start than their them there these they this through under what when where which with would your meet tips example workshop how starting important value people person like love what is the for of and to in on my our who are we you anyone anywhere possible scenario robot event school control navigate button press channels remote teaching member impact lego intro list other choice activities thanks attendance build time now link'.split(' '));
  const words = (deck.keyTerms || []).flatMap((term) => term.match(/[A-Za-z][A-Za-z'-]{2,}/g) || []);
  const qualified = words.filter((term) => (term.length >= 8 || /^[A-Z]{2,}$/.test(term)) && !stopWords.has(term.toLowerCase()) && !people.some((person) => person.name.toLowerCase().includes(term.toLowerCase())));
  return [...new Set(qualified.map((term) => /^[A-Z]{2,}$/.test(term) ? term : term[0].toUpperCase() + term.slice(1).toLowerCase()))].slice(0, 8);
}
async function analyzeWithGemini(deck, apiKey) {
  const deckText = deck.slides.map((slide, index) => `SLIDE ${index + 1}: ${slide.title}\n${slide.text.join('\n')}`).join('\n\n');
  const prompt = `Read every slide and every bulletpoint in the deck below before deciding what matters. Return ONLY valid JSON with this shape: {"mainPoints":["..."],"keyTerms":[{"term":"...","why":"..."}],"people":[{"name":"...","role":"..."}],"quizAnswer":"..."}. Write 8 to 20 specific main points when the deck contains enough material. Combine related bulletpoints into complete explanations, preserve concrete examples, definitions, comparisons, steps, numbers, and pros/cons, and identify which slide each idea comes from when useful. Do not summarize only the first few slides. Exclude personal bios, hobbies, names, navigation instructions, attendance reminders, URLs, and slide labels from main points. Key terms must be actual uncommon, technical, or course-specific vocabulary that a student could be asked to define; do not use ordinary words, generic headings, repeated words, people’s names, or activity labels. If no such terms are supported by the deck, return an empty keyTerms array. List every clearly named person and their role only when supported by the deck. Never invent missing information.\n\nDECK:\n${deckText}`;
  const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(apiKey)}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0.1 } }) });
  if (!response.ok) throw new Error('Gemini could not analyze this deck. Check the API key or use the local analysis.');
  const result = await response.json();
  const text = result.candidates?.[0]?.content?.parts?.[0]?.text || '';
  return JSON.parse(text.replace(/^```json\s*|\s*```$/g, '').trim());
}
function renderQuickAnswers(deck, analysis = null) {
  const lines = contentLines(deck);
  const people = analysis?.people?.length ? analysis.people : extractPeople(deck);
  const summary = analysis?.mainPoints?.length ? analysis.mainPoints : rankPoints(lines.filter(looksLikePoint)).slice(0, 12);
  const terms = analysis?.keyTerms?.length ? analysis.keyTerms.map((item) => typeof item === 'string' ? item : `${item.term}${item.why ? ` — ${item.why}` : ''}`) : extractKeyTerms(deck, people);
  const remember = analysis?.quizAnswer || summary.slice(0, 8).join(' ') || 'Review the Slide text tab for the extracted content.';
  const rememberMarkup = analysis?.quizAnswer ? `<p>${escapeHtml(remember)}</p>` : `<ul>${summary.slice(0, 8).map((point) => `<li>${escapeHtml(point)}</li>`).join('') || '<li>Review the Slide text tab for the extracted content.</li>'}</ul>`;
  const peopleMarkup = people.length ? people.map((person) => `<li><b>${escapeHtml(person.name)}</b> — ${escapeHtml(person.role)}</li>`).join('') : '<li>No named people or roles were clearly identified.</li>';
  const sourceLabel = analysis ? 'AI summary from every bulletpoint' : 'Local context analysis · add a Gemini key for AI summary';
  const sourceGroups = [...new Map(deck.slides.filter((slide) => slide.source).reduce((groups, slide) => { if (!groups.has(slide.source)) groups.set(slide.source, []); groups.get(slide.source).push(slide); return groups; }, new Map())).entries()];
  if (!sourceGroups.length) sourceGroups.push([deck.title, deck.slides]);
  const sourceSummaries = `<div class="source-summaries">${sourceGroups.map(([source, slides]) => { const points = rankPoints(slides.filter((slide) => !/^(intro|thanks|attendance|slide \d+|now:|build time|.*link:?)$/i.test(slide.title)).flatMap((slide) => slide.text.slice(1).filter(looksLikePoint).slice(0, 5).map((point) => `${slide.title}: ${point}`))).slice(0, 36); return `<section><span class="mini-label">SOURCE SUMMARY</span><h3>${escapeHtml(source)}</h3><ul>${points.map((point) => `<li>${escapeHtml(point)}</li>`).join('') || '<li>No complete summary points were found.</li>'}</ul></section>`; }).join('')}</div>`;
  $('#quick-answers').innerHTML = `<div class="quick-heading"><span class="mini-label">SUMMARY + KEY TERMS</span><span>${sourceLabel}</span></div>${sourceSummaries}<div class="answer-grid"><article><span>01 / SUMMARY</span><h3>What are the main points?</h3><ul>${summary.map((point) => `<li>${escapeHtml(point)}</li>`).join('') || '<li>No complete summary points were found.</li>'}</ul></article><article><span>02 / KEY TERMS</span><h3>What words should I know?</h3><ul>${terms.map((term) => `<li>${escapeHtml(term)}</li>`).join('') || '<li>No key terms found.</li>'}</ul></article><article><span>03 / PEOPLE & ROLES</span><h3>Who is mentioned?</h3><ul>${peopleMarkup}</ul></article><article><span>04 / QUIZ CHECK</span><h3>What should I remember?</h3>${rememberMarkup}</article></div>`;
}
function renderTestNotes(deck) {
  const groups = [...new Map(deck.slides.filter((slide) => slide.source).reduce((sourceMap, slide) => { if (!sourceMap.has(slide.source)) sourceMap.set(slide.source, []); sourceMap.get(slide.source).push(slide); return sourceMap; }, new Map())).entries()];
  if (!groups.length) groups.push([deck.title, deck.slides]);
  const sourceNotes = groups.map(([source, slides], sourceIndex) => { const terms = slides.map((slide) => slide.title).filter(Boolean).slice(0, 8); const facts = rankPoints(slides.flatMap((slide) => slide.text.slice(1)).filter((fact) => looksLikePoint(fact))).slice(0, 10); const questions = slides.filter((slide) => !isLowValuePoint(slide.title)).slice(0, 8).map((slide) => `What is the key point of “${slide.title}”?`); return `<section class="test-source"><div class="test-source-header"><span class="mini-label">SOURCE ${String(sourceIndex + 1).padStart(2, '0')}</span><h3>${escapeHtml(source)}</h3></div><div class="notes-grid"><article><span class="notes-label">KEY TERMS</span><ul>${terms.map((term) => `<li>${escapeHtml(term)}</li>`).join('') || '<li>No headings found</li>'}</ul></article><article><span class="notes-label">FACTS TO KNOW</span><ul>${facts.map((fact) => `<li>${escapeHtml(fact)}</li>`).join('') || '<li>No facts found</li>'}</ul></article><article><span class="notes-label">PRACTICE QUESTIONS</span><ul>${questions.map((question) => `<li>${escapeHtml(question)}</li>`).join('') || '<li>No questions generated</li>'}</ul></article></div></section>`; }).join('');
  $('#test-notes').innerHTML = `<div class="test-notes-heading"><div><span class="mini-label">03 / TEST PREP</span><h2 id="test-notes-title">Test-ready notes</h2></div><span>Turn each source into something you can study from</span></div>${sourceNotes}`;
}
function renderOutline(deck) { $('#outline').innerHTML = `<h3>In this deck</h3>${deck.slides.map((slide, index) => `<button class="outline-link" data-slide="${index}" type="button">${String(index + 1).padStart(2, '0')} &nbsp; ${escapeHtml(slide.title.slice(0, 28))}<small>${escapeHtml(slide.source || '')}</small></button>`).join('')}`; }
function sourceHeader(deck, index) { const source = deck.slides[index].source; return source && (index === 0 || deck.slides[index - 1].source !== source) ? `<div class="source-header"><span class="mini-label">SOURCE ${String(deck.sources?.indexOf(source) + 1 || 1).padStart(2, '0')}</span><h3>${escapeHtml(source)}</h3></div>` : ''; }
function studyGuide(deck) { return deck.slides.map((slide, index) => `${sourceHeader(deck, index)}<article class="guide-card"><span class="slide-number">${String(index + 1).padStart(2, '0')} / KEY IDEA${slide.source ? ` · ${escapeHtml(slide.source)}` : ''}</span><h3>${escapeHtml(slide.title)}</h3><p>${escapeHtml(slide.text.slice(1).join(' ') || slide.text[0] || 'No text found on this slide.')}</p>${slide.notes ? `<p><strong>Speaker note:</strong> ${escapeHtml(slide.notes)}</p>` : ''}</article>`).join(''); }
function slideText(deck) { return deck.slides.map((slide, index) => `${sourceHeader(deck, index)}<article class="slide-card" id="slide-${index}"><span class="slide-number">SLIDE ${String(index + 1).padStart(2, '0')}${slide.source ? ` · ${escapeHtml(slide.source)}` : ''}</span><h3>${escapeHtml(slide.title)}</h3><ul>${slide.text.map((line) => `<li>${escapeHtml(line)}</li>`).join('') || '<li>No text found.</li>'}</ul></article>`).join(''); }
function linksMedia(deck) { const items = [...deck.links.map((url) => [url, 'Link']), ...deck.images.map((url) => [url, 'Image'])]; return items.length ? items.map(([url, type]) => `<div class="link-card"><a href="${escapeHtml(url)}" target="_blank" rel="noreferrer">${escapeHtml(url)}</a><span>${type}</span></div>`).join('') : '<p class="empty-state">No links or image references were found in this deck.</p>'; }
function renderView() { const deck = state.deck; $('#content-view').innerHTML = state.view === 'guide' ? studyGuide(deck) : state.view === 'slides' ? slideText(deck) : linksMedia(deck); document.querySelectorAll('.tab').forEach((tab) => tab.classList.toggle('is-active', tab.dataset.view === state.view)); }
function renderDeck(deck) { state.deck = deck; $('#deck-title').textContent = deck.title; $('#deck-subtitle').textContent = `${deck.slides.length} slides mined into assignment-ready notes`; renderQuickAnswers(deck); renderStats(deck); renderOutline(deck); renderView(); renderTestNotes(deck); $('#results').classList.remove('is-hidden'); $('#results').scrollIntoView({ behavior: 'smooth', block: 'start' }); const aiKey = $('#ai-key').value.trim(); if (aiKey) analyzeWithGemini(deck, aiKey).then((analysis) => { if (state.deck === deck) renderQuickAnswers(deck, analysis); }).catch((error) => setError(`${error.message} The local analysis is still shown.`)); }

function markdown() { const deck = state.deck; const facts = deck.slides.flatMap((slide) => slide.text.slice(1)).filter(Boolean).slice(0, 6); const terms = deck.slides.map((slide) => slide.title).filter(Boolean).slice(0, 6); return `# ${deck.title}\n\n## Study guide\n\n${deck.slides.map((slide, index) => `### ${index + 1}. ${slide.title}\n\n${slide.text.slice(1).join(' ') || slide.text[0] || 'No text found.'}${slide.notes ? `\n\n**Speaker note:** ${slide.notes}` : ''}`).join('\n\n')}\n\n## Test-ready notes\n\n### Key terms\n\n${terms.map((term) => `- ${term}`).join('\n') || '- None found'}\n\n### Facts to know\n\n${facts.map((fact) => `- ${fact}`).join('\n') || '- None found'}\n\n### Practice questions\n\n${terms.map((term) => `- What is the key point of “${term}”?`).join('\n') || '- None found'}\n\n## Links and media\n\n${[...deck.links, ...deck.images].map((item) => `- ${item}`).join('\n') || '- None found'}`; }

$('#mine-button').addEventListener('click', async () => { clearError(); setLoading(true); try { renderDeck(await fetchDeck()); } catch (error) { setError(error.message); } finally { setLoading(false); } });
$('#demo-button').addEventListener('click', () => { document.querySelector('.slides-url').value = 'sample://attention-economy'; clearError(); renderDeck(SAMPLE); });
async function processFiles(files) { if (!files?.length) return; clearError(); setLoading(true); try { renderDeck(await parseUploadedFiles(files)); } catch (error) { setError(error.message); } finally { setLoading(false); } }
$('#deck-file').addEventListener('change', (event) => processFiles(event.target.files));
$('#drop-zone').addEventListener('dragover', (event) => { event.preventDefault(); $('#drop-zone').classList.add('is-dragging'); });
$('#drop-zone').addEventListener('dragleave', (event) => { if (!event.relatedTarget || !$('#drop-zone').contains(event.relatedTarget)) $('#drop-zone').classList.remove('is-dragging'); });
$('#drop-zone').addEventListener('drop', (event) => { event.preventDefault(); $('#drop-zone').classList.remove('is-dragging'); processFiles(event.dataTransfer.files); });
$('#add-url').addEventListener('click', () => { const entry = document.createElement('div'); entry.className = 'url-entry'; entry.innerHTML = '<input class="slides-url" type="url" placeholder="https://docs.google.com/presentation/d/..." autocomplete="off"><button class="remove-url" type="button" aria-label="Clear link">×</button>'; $('#url-list').appendChild(entry); entry.querySelector('input').focus(); });
$('#url-list').addEventListener('click', (event) => { const button = event.target.closest('.remove-url'); if (!button) return; button.closest('.url-entry').querySelector('.slides-url').value = ''; button.closest('.url-entry').querySelector('.slides-url').focus(); });
document.querySelectorAll('.tab').forEach((tab) => tab.addEventListener('click', () => { state.view = tab.dataset.view; renderView(); }));
$('#outline').addEventListener('click', (event) => {
  const button = event.target.closest('[data-slide]');
  if (!button) return;
  state.view = 'slides';
  renderView();
  document.querySelector(`#slide-${button.dataset.slide}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
});
$('#copy-button').addEventListener('click', async () => { await navigator.clipboard.writeText(markdown()); $('#copy-button').textContent = 'Copied'; setTimeout(() => { $('#copy-button').textContent = 'Copy study guide'; }, 1400); });
$('#download-button').addEventListener('click', () => { const blob = new Blob([markdown()], { type: 'text/markdown' }); const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = `${state.deck.title.replace(/[^a-z0-9]+/gi, '-').toLowerCase() || 'slides'}-study-guide.md`; link.click(); URL.revokeObjectURL(link.href); });
document.addEventListener('keydown', (event) => { if (event.target.matches('.slides-url') && event.key === 'Enter') { event.preventDefault(); $('#mine-button').click(); } });

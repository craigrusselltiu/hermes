// Tauri API imports
const { invoke } = window.__TAURI__.core;
const { open } = window.__TAURI__.dialog;
const { listen } = window.__TAURI__.event;

// DOM refs
const welcomeScreen = document.getElementById('welcome-screen');
const documentView = document.getElementById('document-view');
const desk = document.getElementById('desk');
const commentsPanel = document.getElementById('comments-panel');
const commentsList = document.getElementById('comments-list');
const fileInfo = document.getElementById('file-info');
const themeIcon = document.getElementById('theme-icon');
const findBar = document.getElementById('find-bar');
const findInput = document.getElementById('find-input');
const findCount = document.getElementById('find-count');

// State
let currentDocument = null;
let currentFilePath = null;
let commentsVisible = false;
let findMatches = [];
let findIndex = -1;

// --- Initialization ---

document.addEventListener('DOMContentLoaded', () => {
    initializeTheme();
    setupEventListeners();
    setupKeyboardShortcuts();
    setupTauriListeners();
    setupDragAndDrop();
});

function setupEventListeners() {
    document.getElementById('open-file-btn').addEventListener('click', handleOpenFile);
    document.getElementById('welcome-open-btn').addEventListener('click', handleOpenFile);
    document.getElementById('theme-toggle-btn').addEventListener('click', toggleTheme);
    document.getElementById('toggle-comments-btn').addEventListener('click', toggleComments);
    document.getElementById('close-comments-btn').addEventListener('click', toggleComments);
    document.getElementById('find-close-btn').addEventListener('click', closeFindBar);
    document.getElementById('find-next-btn').addEventListener('click', () => navigateFind(1));
    document.getElementById('find-prev-btn').addEventListener('click', () => navigateFind(-1));
    findInput.addEventListener('input', performFind);
    findInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') navigateFind(e.shiftKey ? -1 : 1);
        if (e.key === 'Escape') closeFindBar();
    });
}

function setupKeyboardShortcuts() {
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'o') { e.preventDefault(); handleOpenFile(); }
        if (e.ctrlKey && e.key === 'd') { e.preventDefault(); toggleTheme(); }
        if (e.ctrlKey && e.key === ']') { e.preventDefault(); toggleComments(); }
        if (e.ctrlKey && e.key === 'f') { e.preventDefault(); openFindBar(); }
        if (e.ctrlKey && e.key === 'q') { e.preventDefault(); window.__TAURI__?.process?.exit(0); }
    });
}

async function setupTauriListeners() {
    try {
        await listen('tauri://drag-drop', async (event) => {
            const paths = event.payload?.paths || [];
            const docxPath = paths.find(p => p.toLowerCase().endsWith('.docx'));
            if (docxPath) {
                await loadDocument(docxPath);
            }
        });
    } catch (err) {
        console.log('Drag-drop listener setup:', err);
    }
}

function setupDragAndDrop() {
    document.addEventListener('dragenter', (e) => { e.preventDefault(); document.body.classList.add('drag-over'); });
    document.addEventListener('dragover', (e) => { e.preventDefault(); });
    document.addEventListener('dragleave', (e) => { e.preventDefault(); document.body.classList.remove('drag-over'); });
    document.addEventListener('drop', (e) => { e.preventDefault(); document.body.classList.remove('drag-over'); });
}

// --- File handling ---

async function handleOpenFile() {
    try {
        const selected = await open({
            multiple: false,
            filters: [{ name: 'Word Documents', extensions: ['docx'] }]
        });
        if (selected) {
            const path = selected.path || selected;
            await loadDocument(path);
        }
    } catch (err) {
        showError('Could not open file: ' + err);
    }
}

async function loadDocument(path) {
    currentFilePath = path;
    fileInfo.textContent = 'Loading...';

    try {
        const doc = await invoke('open_docx', { path });
        currentDocument = doc;
        fileInfo.textContent = '';
        renderDocument(doc);
    } catch (err) {
        showError(err);
    }
}

function showError(msg) {
    fileInfo.textContent = 'Error: ' + msg;
    fileInfo.style.color = '#e74c3c';
    setTimeout(() => { fileInfo.style.color = ''; }, 5000);
}

// --- Document rendering ---

function renderDocument(doc) {
    welcomeScreen.style.display = 'none';
    documentView.style.display = 'flex';
    desk.innerHTML = '';

    // Build pages split by PageBreak
    const pages = splitIntoPages(doc.body);

    pages.forEach((pageBlocks, pageIdx) => {
        const page = document.createElement('div');
        page.className = 'page';

        // Header
        if (doc.headers && doc.headers.length > 0) {
            const header = document.createElement('div');
            header.className = 'page-header';
            renderBlocks(doc.headers[0].content, header, doc);
            page.appendChild(header);
        }

        // Body content
        const content = document.createElement('div');
        content.className = 'page-content';
        renderBlocks(pageBlocks, content, doc);
        page.appendChild(content);

        // Footnotes for this page
        const pageFootnotes = collectFootnotes(pageBlocks, doc.footnotes);
        if (pageFootnotes.length > 0) {
            const fnSection = document.createElement('div');
            fnSection.className = 'page-footnotes';
            fnSection.appendChild(document.createElement('hr'));
            pageFootnotes.forEach(fn => {
                const fnEl = document.createElement('div');
                fnEl.className = 'footnote';
                fnEl.id = 'footnote-' + fn.id;
                const marker = document.createElement('sup');
                marker.textContent = fn.id;
                fnEl.appendChild(marker);
                const fnContent = document.createElement('span');
                renderBlocks(fn.content, fnContent, doc);
                fnEl.appendChild(fnContent);
                fnSection.appendChild(fnEl);
            });
            page.appendChild(fnSection);
        }

        // Footer
        if (doc.footers && doc.footers.length > 0) {
            const footer = document.createElement('div');
            footer.className = 'page-footer';
            renderBlocks(doc.footers[0].content, footer, doc);
            page.appendChild(footer);
        }

        desk.appendChild(page);
    });

    // Render comments panel
    renderComments(doc.comments);

    // Auto-show comments panel if there are comments
    if (doc.comments && doc.comments.length > 0) {
        commentsVisible = true;
        commentsPanel.style.display = 'flex';
    }

    // Update window title
    const filename = currentFilePath.split(/[\\/]/).pop();
    document.title = filename + ' - Hermes';
}

function splitIntoPages(blocks) {
    const pages = [];
    let current = [];

    for (const block of blocks) {
        if (block.type === 'page_break') {
            pages.push(current);
            current = [];
        } else {
            current.push(block);
        }
    }
    if (current.length > 0) {
        pages.push(current);
    }
    if (pages.length === 0) {
        pages.push([]);
    }
    return pages;
}

function renderBlocks(blocks, container, doc) {
    if (!blocks) return;

    for (const block of blocks) {
        switch (block.type) {
            case 'paragraph':
                renderParagraph(block, container, doc);
                break;
            case 'table':
                renderTable(block, container, doc);
                break;
            case 'page_break':
                break;
        }
    }
}

function renderParagraph(para, container, doc) {
    // Determine if heading
    const style = para.style ? (doc.styles?.[para.style] || null) : null;
    const headingLevel = style?.heading_level || detectHeadingLevel(para.style);

    let el;
    if (headingLevel && headingLevel >= 1 && headingLevel <= 6) {
        el = document.createElement('h' + headingLevel);
    } else {
        el = document.createElement('p');
    }

    // Alignment
    const align = para.alignment || style?.alignment;
    if (align) {
        el.style.textAlign = align;
    }

    // Render runs
    if (para.runs) {
        for (const run of para.runs) {
            renderRun(run, el, doc);
        }
    }

    // Empty paragraph - add non-breaking space for spacing
    if (el.childNodes.length === 0) {
        el.innerHTML = '&nbsp;';
    }

    container.appendChild(el);
}

function renderRun(run, container, doc) {
    // Image run
    if (run.image_id && doc.images && doc.images[run.image_id]) {
        const img = document.createElement('img');
        img.src = doc.images[run.image_id];
        img.className = 'doc-image';
        img.alt = 'Document image';
        container.appendChild(img);
        return;
    }

    // Footnote reference
    if (run.footnote_ref) {
        const sup = document.createElement('sup');
        const link = document.createElement('a');
        link.href = '#footnote-' + run.footnote_ref;
        link.className = 'footnote-ref';
        link.textContent = run.footnote_ref;
        sup.appendChild(link);
        container.appendChild(sup);
    }

    // Text run
    if (!run.text || run.text.length === 0) return;

    const span = document.createElement('span');
    span.textContent = run.text;

    // Apply formatting
    if (run.bold) span.style.fontWeight = 'bold';
    if (run.italic) span.style.fontStyle = 'italic';
    if (run.underline) span.style.textDecoration = 'underline';
    if (run.strikethrough) {
        span.style.textDecoration = (span.style.textDecoration || '') +
            (span.style.textDecoration ? ' line-through' : 'line-through');
    }
    if (run.font_size) span.style.fontSize = run.font_size + 'pt';
    if (run.font_family) span.style.fontFamily = run.font_family;
    if (run.color && run.color !== 'auto') span.style.color = '#' + run.color;
    if (run.highlight) {
        span.style.backgroundColor = highlightColorMap[run.highlight] || run.highlight;
    }

    // Comment reference
    if (run.comment_ref != null) {
        span.classList.add('commented-text');
        span.dataset.commentId = run.comment_ref;
        span.addEventListener('click', () => scrollToComment(run.comment_ref));
    }

    container.appendChild(span);
}

function renderTable(table, container, doc) {
    const tableEl = document.createElement('table');
    tableEl.className = 'doc-table';

    if (table.rows) {
        for (const row of table.rows) {
            const tr = document.createElement('tr');
            if (row.cells) {
                for (const cell of row.cells) {
                    const td = document.createElement('td');
                    if (cell.col_span > 1) td.colSpan = cell.col_span;
                    if (cell.row_span > 1) td.rowSpan = cell.row_span;
                    if (cell.shading) td.style.backgroundColor = cell.shading;
                    renderBlocks(cell.content, td, doc);
                    tr.appendChild(td);
                }
            }
            tableEl.appendChild(tr);
        }
    }

    container.appendChild(tableEl);
}

function collectFootnotes(blocks, footnotes) {
    if (!footnotes || !blocks) return [];
    const refs = new Set();
    const collectRefs = (blocks) => {
        for (const b of blocks) {
            if (b.type === 'paragraph' && b.runs) {
                for (const r of b.runs) {
                    if (r.footnote_ref) refs.add(r.footnote_ref);
                }
            }
            if (b.type === 'table' && b.rows) {
                for (const row of b.rows) {
                    for (const cell of row.cells || []) {
                        collectRefs(cell.content || []);
                    }
                }
            }
        }
    };
    collectRefs(blocks);
    return footnotes.filter(fn => refs.has(fn.id));
}

function detectHeadingLevel(styleName) {
    if (!styleName) return 0;
    const lower = styleName.toLowerCase();
    const match = lower.match(/heading\s*(\d)/);
    if (match) return parseInt(match[1]);
    if (lower === 'title') return 1;
    if (lower === 'subtitle') return 2;
    return 0;
}

const highlightColorMap = {
    yellow: '#ffff00',
    green: '#00ff00',
    cyan: '#00ffff',
    magenta: '#ff00ff',
    blue: '#0000ff',
    red: '#ff0000',
    darkBlue: '#00008b',
    darkCyan: '#008b8b',
    darkGreen: '#006400',
    darkMagenta: '#8b008b',
    darkRed: '#8b0000',
    darkYellow: '#9b870c',
    darkGray: '#a9a9a9',
    lightGray: '#d3d3d3',
    black: '#000000',
};

// --- Comments ---

function renderComments(comments) {
    commentsList.innerHTML = '';
    if (!comments || comments.length === 0) {
        commentsList.innerHTML = '<p class="no-comments">No comments in this document.</p>';
        return;
    }

    for (const comment of comments) {
        const el = document.createElement('div');
        el.className = 'comment-card';
        el.id = 'comment-' + comment.id;
        el.innerHTML = `
            <div class="comment-meta">
                <strong>${escapeHtml(comment.author)}</strong>
                ${comment.date ? '<span class="comment-date">' + formatDate(comment.date) + '</span>' : ''}
            </div>
            <div class="comment-text">${escapeHtml(comment.text)}</div>
        `;
        el.addEventListener('click', () => scrollToCommentedText(comment.id));
        commentsList.appendChild(el);
    }
}

function toggleComments() {
    commentsVisible = !commentsVisible;
    commentsPanel.style.display = commentsVisible ? 'flex' : 'none';
}

function scrollToComment(commentId) {
    const el = document.getElementById('comment-' + commentId);
    if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.classList.add('highlight');
        setTimeout(() => el.classList.remove('highlight'), 2000);
    }
    if (!commentsVisible) toggleComments();
}

function scrollToCommentedText(commentId) {
    const el = document.querySelector(`[data-comment-id="${commentId}"]`);
    if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.classList.add('flash');
        setTimeout(() => el.classList.remove('flash'), 2000);
    }
}

// --- Find ---

function openFindBar() {
    findBar.style.display = 'flex';
    findInput.focus();
    findInput.select();
}

function closeFindBar() {
    findBar.style.display = 'none';
    clearFindHighlights();
    findMatches = [];
    findIndex = -1;
    findCount.textContent = '';
}

function performFind() {
    clearFindHighlights();
    const query = findInput.value.trim().toLowerCase();
    if (!query) { findCount.textContent = ''; findMatches = []; return; }

    findMatches = [];
    const walker = document.createTreeWalker(desk, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
        const node = walker.currentNode;
        const idx = node.textContent.toLowerCase().indexOf(query);
        if (idx >= 0) {
            const range = document.createRange();
            range.setStart(node, idx);
            range.setEnd(node, idx + query.length);
            const mark = document.createElement('mark');
            mark.className = 'find-highlight';
            range.surroundContents(mark);
            findMatches.push(mark);
        }
    }

    findCount.textContent = findMatches.length + ' found';
    findIndex = findMatches.length > 0 ? 0 : -1;
    if (findIndex >= 0) highlightCurrentMatch();
}

function navigateFind(dir) {
    if (findMatches.length === 0) return;
    findIndex = (findIndex + dir + findMatches.length) % findMatches.length;
    highlightCurrentMatch();
}

function highlightCurrentMatch() {
    findMatches.forEach(m => m.classList.remove('current'));
    if (findIndex >= 0 && findIndex < findMatches.length) {
        findMatches[findIndex].classList.add('current');
        findMatches[findIndex].scrollIntoView({ behavior: 'smooth', block: 'center' });
        findCount.textContent = (findIndex + 1) + '/' + findMatches.length;
    }
}

function clearFindHighlights() {
    document.querySelectorAll('mark.find-highlight').forEach(mark => {
        const parent = mark.parentNode;
        parent.replaceChild(document.createTextNode(mark.textContent), mark);
        parent.normalize();
    });
}

// --- Theme ---

function initializeTheme() {
    const saved = localStorage.getItem('hermes-theme') || 'light';
    setTheme(saved);
}

function toggleTheme() {
    const current = document.body.classList.contains('dark-theme') ? 'dark' : 'light';
    setTheme(current === 'light' ? 'dark' : 'light');
}

function setTheme(theme) {
    if (theme === 'dark') {
        document.body.classList.add('dark-theme');
        themeIcon.innerHTML = '&#9788;';
    } else {
        document.body.classList.remove('dark-theme');
        themeIcon.innerHTML = '&#9789;';
    }
    localStorage.setItem('hermes-theme', theme);
}

// --- Utilities ---

function escapeHtml(text) {
    const d = document.createElement('div');
    d.textContent = text;
    return d.innerHTML;
}

function formatDate(dateStr) {
    try {
        return new Date(dateStr).toLocaleDateString(undefined, {
            year: 'numeric', month: 'short', day: 'numeric'
        });
    } catch { return dateStr; }
}

function getFileName(path) {
    return path ? path.split(/[\\/]/).pop() : '';
}
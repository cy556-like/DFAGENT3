const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = process.env.RELATED_SOURCE || path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'app/static/js/app.js'), 'utf8');
const prefix = source.match(/`(qzagent3|sbagent|dfagent3):\$\{encodeURIComponent/)[1];
const repo = prefix.toUpperCase();

function extract(name) {
    const re = new RegExp(`^(?:async )?function ${name}\\([^\\n]*\\) \\{`, 'm');
    const match = re.exec(source);
    assert.ok(match, name);
    const end = source.indexOf('\n}', match.index + match[0].length);
    return source.slice(match.index, end + 2);
}

function deferred() {
    let resolve;
    const promise = new Promise(done => { resolve = done; });
    return { promise, resolve };
}

function sandbox(extraNames = []) {
    const storage = new Map();
    const elements = new Map();
    const makeElement = () => ({
        value: '', disabled: false, innerHTML: '', textContent: '', style: {}, options: [], selectedIndex: 0,
        classList: { add() {}, remove() {}, toggle() {} }, appendChild(option) { this.options.push(option); },
        querySelectorAll() { return []; }, querySelector() { return null; }, replaceChildren() { this.innerHTML = ''; },
    });
    let body = '';
    const context = vm.createContext({
        console: { warn() {}, error() {}, log() {} },
        localStorage: { getItem: key => storage.get(key) ?? null, setItem: (key, value) => storage.set(key, value), removeItem: key => storage.delete(key) },
        document: { getElementById(id) { if (!elements.has(id)) elements.set(id, makeElement()); return elements.get(id); },
            createElement: makeElement, querySelector: () => null, querySelectorAll: () => [],
            body: { classList: { add() {}, remove() {} } }, hidden: false },
        setTimeout: () => 1, clearTimeout() {}, clearInterval() {}, Object, JSON, Map, Set, DOMException,
        history: { state: null, replaceState() {} },
        loadAgentActiveChatIds() {}, showToast() {}, restoreSelectedSkillForCurrentUser() {},
        updateHeaderKbVisibility() {}, resetStreamingUI() {}, updateCenteredMode() {},
        showQualityWorkspaceChrome() {}, isQualityImprovementSubagent: () => false,
        restoreCachedGenerationForChat() {}, restoreManualGenerationForChat: async () => {},
        addMessageToUI(role, text) { body += text; }, scrollToBottom() {}, renderMyAgents() {},
        clearNavigationState() {}, sessionStorage: { removeItem() {} },
        apiHeaders: () => ({}), readApiResponse: async response => response.json(),
        filterAgents: agents => Array.isArray(agents) ? agents : [{ id: 'built-in', chat_ids: [] }],
        apiFetch: async () => { throw new Error('fetch not set'); },
        manualJobIsActive: job => job?.status === 'running',
        getBody: () => body,
        AbortController,
    });
    vm.runInContext(`
        let currentUser = 'admin', authToken = 'token-a', userRole = 'admin', authSource = 'web';
        let accountSessionVersion = 1, authRequestVersion = 0, accountCacheWritesPaused = false;
        const accountRequestVersions = Object.create(null);
        let currentChatId = 'a-1', currentAgentId = 'built-in', currentMode = 'agent';
        let allChats = [], myAgents = [], modeChatId = {}, agentActiveChatId = {};
        let currentAbortController = null, thinkingInterval = null, isLoading = false;
        let selectedFile = null, selectedFileBase64 = null, selectedSkill = null, lastMessageText = '';
        let _lastSyncedAgentsHash = '', _syncAgentsLock = false, _syncAgentsLastTime = 0;
        const _SYNC_AGENTS_COOLDOWN = 5000;
        let _chatListLoadSeq = 0, _chatHistoryLoadSeq = 0, _kbDocsLoadSeq = 0, _chatCreateSeq = 0;
        let currentWorkspaceId = null, currentQualityWorkspaceId = null, surveyReturnChatId = null;
        let surveyUploadedFiles = [];
        let modelSwitchInProgress = false, currentModelId = 'auto', fullKnowledgeAdmin = false;
        let agentKbUploadMode = false, chatHistoryLoadSeq = 0;
        const manualGenerationJobs = new Map(), manualGenerationPollers = new Map();
    `, context);
    for (const name of ['captureAccountContext', 'isCurrentAccount', 'accountCacheKey', 'readAccountStorage', 'writeAccountStorage', 'resetAccountState', 'beginAccountSession', ...extraNames]) {
        vm.runInContext(extract(name), context);
    }
    context.storage = storage;
    context.elements = elements;
    return context;
}

function switchAccount(ctx, username = 'adminquanzhi', token = 'token-b') {
    vm.runInContext(`resetAccountState(); currentUser=${JSON.stringify(username)}; authToken=${JSON.stringify(token)}; beginAccountSession();`, ctx);
}

test('account cache never shares admin and adminquanzhi entries', () => {
    const ctx = sandbox();
    vm.runInContext(`writeAccountStorage('agents', 'private-a');`, ctx);
    switchAccount(ctx);
    assert.equal(vm.runInContext(`readAccountStorage('agents')`, ctx), null);
    vm.runInContext(`writeAccountStorage('agents', 'private-b');`, ctx);
    assert.equal(ctx.storage.get(`${repo.toLowerCase()}:admin:agents`), 'private-a');
    assert.equal(ctx.storage.get(`${repo.toLowerCase()}:adminquanzhi:agents`), 'private-b');
});

test('ownerless legacy cache is not read', () => {
    const ctx = sandbox();
    ctx.storage.set('forgeAgents', JSON.stringify([{ id: 'foreign', chat_ids: ['other-chat'] }]));
    ctx.storage.set('quanzhiAgents', ctx.storage.get('forgeAgents'));
    vm.runInContext('beginAccountSession()', ctx);
    assert.equal(vm.runInContext(`myAgents[0].id`, ctx), 'built-in');
});

test('new login keeps its own cached custom agents', () => {
    const ctx = sandbox();
    ctx.storage.set(`${repo.toLowerCase()}:admin:agents`, JSON.stringify([{ id: 'custom', chat_ids: ['a-1'] }]));
    vm.runInContext('beginAccountSession()', ctx);
    assert.equal(vm.runInContext('myAgents[0].id', ctx), 'custom');
});

test('old response rejected after logout and same username/token login', () => {
    const ctx = sandbox();
    vm.runInContext('globalThis.old = captureAccountContext()', ctx);
    switchAccount(ctx, 'admin', 'token-a');
    assert.equal(vm.runInContext('isCurrentAccount(old)', ctx), false);
});

test('account reset clears private state and aborts old stream', () => {
    const ctx = sandbox();
    let aborted = false;
    ctx.abort = () => { aborted = true; };
    vm.runInContext(`currentAbortController={abort}; allChats=[{chat_id:'a-1'}]; selectedFile={name:'a.txt'}; selectedFileBase64='secret'; resetAccountState();`, ctx);
    assert.ok(aborted);
    for (const expression of ['allChats.length === 0', 'selectedFile === null', 'selectedFileBase64 === null', 'currentChatId === null', 'currentAgentId === null']) {
        assert.equal(vm.runInContext(expression, ctx), true, expression);
    }
});

test('delayed model list cannot overwrite new account selector', async () => {
    const ctx = sandbox(['loadModels']);
    const request = deferred();
    ctx.fetch = ctx.apiFetch = () => request.promise;
    const pending = vm.runInContext('loadModels()', ctx);
    switchAccount(ctx);
    ctx.elements.get('modelSelect').value = 'qwen3.7-plus';
    request.resolve({ ok: true, json: async () => ({ models: [{ id: 'glm-5.3', name: 'foreign' }], current: 'glm-5.3' }) });
    await pending;
    assert.equal(ctx.elements.get('modelSelect').value, 'qwen3.7-plus');
});

test('delayed model failure cannot replace new account selector', async () => {
    const ctx = sandbox(['loadModels']);
    const request = deferred();
    ctx.fetch = ctx.apiFetch = () => request.promise;
    const pending = vm.runInContext('loadModels()', ctx);
    switchAccount(ctx);
    ctx.elements.get('modelSelect').value = 'qwen3.7-plus';
    request.resolve({ ok: false, status: 401, json: async () => ({}) });
    await pending;
    assert.equal(ctx.elements.get('modelSelect').value, 'qwen3.7-plus');
});

test('delayed history cannot render previous account content', async () => {
    const ctx = sandbox(['loadChatHistory']);
    const request = deferred();
    ctx.fetch = ctx.apiFetch = () => request.promise;
    const pending = vm.runInContext(`loadChatHistory('a-1')`, ctx);
    switchAccount(ctx);
    request.resolve({ ok: true, status: 200, json: async () => ({ messages: [{ role: 'assistant', content: 'private-a' }] }) });
    await pending;
    assert.equal(ctx.getBody(), '');
});

test('model refresh honors saved account selection', async () => {
    const ctx = sandbox(['loadModels']);
    ctx.fetch = ctx.apiFetch = async () => ({ ok: true, json: async () => ({ current: 'glm-5.3', models: [{ id: 'glm-5.3', name: 'GLM' }, { id: 'auto-model', name: 'Auto' }] }) });
    await vm.runInContext('loadModels()', ctx);
    if (repo === 'DFAGENT3') {
        assert.equal(ctx.elements.get('modelSelect').options[0].selected, true);
    } else {
        assert.equal(ctx.elements.get('modelSelect').value, 'glm-5.3');
    }
});

test('saved list removes foreign and deleted chat IDs', async () => {
    const ctx = sandbox(['rebuildChatIdsFromServer']);
    const chats = [{ chat_id: 'live', agent_id: 'built-in' }];
    ctx.fetch = async () => ({ json: async () => ({ success: true, chats }) });
    ctx.fetchAllChatPages = async () => ({ success: true, chats });
    ctx.fetchAllChatsForUser = async () => chats;
    vm.runInContext(`myAgents=[{id:'built-in',chat_ids:['foreign','deleted']}];`, ctx);
    await vm.runInContext('rebuildChatIdsFromServer()', ctx);
    assert.deepEqual(Array.from(vm.runInContext('myAgents[0].chat_ids', ctx)), ['live']);
});

test('a duplicate debounced sync never invalidates the running sync or leaves its lock stuck', async () => {
    const ctx = sandbox(['syncAgentsFromServer']);
    const request = deferred();
    ctx.fetch = () => request.promise;
    ctx.rebuildChatIdsFromServer = async () => {};
    const first = vm.runInContext('syncAgentsFromServer(true)', ctx);
    await vm.runInContext('syncAgentsFromServer()', ctx);
    request.resolve({ json: async () => ({ success: true, agents: [] }) });
    await first;
    assert.equal(vm.runInContext('_syncAgentsLock', ctx), false);
});

if (repo === 'QZAGENT3') {
    test('late survey load never fills or caches a previous account form', async () => {
        const ctx = sandbox(['loadSurveyData']);
        const request = deferred();
        ctx.apiFetch = () => request.promise;
        let filled = false;
        ctx._fillSurveyFormUI = () => { filled = true; };
        const pending = vm.runInContext('loadSurveyData()', ctx);
        switchAccount(ctx);
        request.resolve({ json: async () => ({ success: true, survey_data: { sv_company_name: 'private-a' } }) });
        await pending;
        assert.equal(filled, false);
        assert.equal(vm.runInContext(`readAccountStorage('surveyData')`, ctx), null);
    });
    test('stale 401 cannot log out newly authenticated account', async () => {
        const ctx = sandbox(['apiFetch']);
        const request = deferred();
        ctx.fetch = () => request.promise;
        let logout = false;
        ctx.doLogout = () => { logout = true; };
        const pending = vm.runInContext(`apiFetch('/api/v1/history/a-1')`, ctx);
        switchAccount(ctx);
        request.resolve({ status: 401 });
        await assert.rejects(pending, { name: 'AbortError' });
        assert.equal(logout, false);
    });

    test('generation progress polling captures its owning account', () => {
        assert.match(extract('scheduleManualJobPoll'), /captureAccountContext/);
        assert.match(extract('restoreManualGenerationForChat'), /isCurrentAccount/);
        assert.match(extract('startManualGenerationJob'), /isCurrentAccount/);
    });
}

if (repo !== 'DFAGENT3') {
    for (const fails of [false, true]) {
        test(`late SSO logout ${fails ? 'failure' : 'response'} cannot clear a new login`, async () => {
            const ctx = sandbox(['doLogout']);
            const request = deferred();
            ctx.fetch = () => request.promise;
            ctx.stopGeneration = () => {};
            vm.runInContext(`authSource=${JSON.stringify(repo === 'QZAGENT3' ? 'jsxf_demo' : 'feishu')}`, ctx);
            const pending = vm.runInContext('doLogout()', ctx);
            switchAccount(ctx);
            request.resolve(fails ? Promise.reject(new Error('offline')) : { ok: true });
            await pending;
            assert.equal(vm.runInContext('currentUser', ctx), 'adminquanzhi');
            assert.equal(vm.runInContext('authToken', ctx), 'token-b');
        });
    }
}

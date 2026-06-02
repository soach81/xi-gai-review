"""Add GitHub Gist sync functionality to index.html."""
import re

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

# ============================================================
# 1. ADD CSS for login modal and sync status
# ============================================================
sync_css = """
/* Sync & Login */
.sync-status { font-size: 11px; color: var(--text2); cursor: pointer; padding: 4px 10px;
  border-radius: 16px; border: 1px solid var(--border); background: var(--bg);
  display: flex; align-items: center; gap: 4px; white-space: nowrap; }
.sync-status .dot { width: 7px; height: 7px; border-radius: 50%; background: var(--text2); flex-shrink: 0; }
.sync-status .dot.online { background: var(--success); }
.sync-status .dot.syncing { background: var(--warn); animation: pulse 1s infinite; }
@keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.4; } }

.modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 500;
  display: flex; align-items: center; justify-content: center; }
.modal-box { background: var(--surface); border-radius: var(--radius); padding: 28px;
  max-width: 420px; width: 90vw; box-shadow: var(--shadow-lg); }
.modal-box h3 { font-size: 17px; font-weight: 700; margin-bottom: 16px; }
.modal-box label { font-size: 13px; font-weight: 600; display: block; margin: 10px 0 4px; color: var(--text2); }
.modal-box input { width: 100%; padding: 10px 12px; border: 1px solid var(--border);
  border-radius: var(--radius-sm); font-size: 14px; background: var(--bg); color: var(--text); }
.modal-box .modal-actions { display: flex; gap: 10px; margin-top: 18px; justify-content: flex-end; }
.modal-box .modal-actions button { padding: 8px 20px; border-radius: 20px; font-size: 14px; cursor: pointer; border: 1px solid var(--border); background: var(--bg); color: var(--text); }
.modal-box .modal-actions .btn-primary { background: var(--gradient); color: #fff; border: none; font-weight: 600; }
.modal-box .hint { font-size: 11px; color: var(--text2); margin-top: 6px; line-height: 1.5; }
.modal-box .hint a { color: var(--primary); }
.modal-box .error { font-size: 12px; color: var(--danger); margin-top: 6px; display: none; }
"""

# Insert CSS before </style>
html = html.replace('</style>', sync_css + '\n</style>')

# ============================================================
# 2. ADD login modal HTML and sync status to topbar
# ============================================================

# Add sync status button after user-select in topbar
old_topbar = '<span class="spacer"></span>'
new_topbar = '<span class="spacer"></span>\n  <span class="sync-status" id="sync-status" onclick="showSyncModal()" title="点击登录以跨设备同步"><span class="dot" id="sync-dot"></span><span id="sync-label">未登录</span></span>'
html = html.replace(old_topbar, new_topbar, 1)

# Add login modal div before </body>
modal_html = """
<!-- Login Modal -->
<div class="modal-overlay" id="login-modal" style="display:none" onclick="if(event.target===this)this.style.display='none'">
  <div class="modal-box">
    <h3 id="login-modal-title">登录以同步数据</h3>
    <label>GitHub 用户名</label>
    <input id="login-user" placeholder="your-username" autocomplete="username">
    <label>Personal Access Token</label>
    <input id="login-token" type="password" placeholder="ghp_xxxx..." autocomplete="off">
    <div class="error" id="login-error"></div>
    <div class="hint">需要 <a href="https://github.com/settings/tokens/new?scopes=gist&description=xi-gai-sync" target="_blank">创建 PAT</a>，勾选 <b>gist</b> 权限即可。<br>数据存在你的 Secret Gist 里，不会公开。</div>
    <div class="modal-actions">
      <button onclick="document.getElementById('login-modal').style.display='none'">取消</button>
      <button class="btn-primary" onclick="doLogin()">登录</button>
    </div>
  </div>
</div>
"""
html = html.replace('</body>', modal_html + '\n</body>')

# ============================================================
# 3. ADD SyncManager JS and modify saveJSON
# ============================================================

# Find the end of the init() function call and add SyncManager after it
sync_js = r"""

// ====== SYNC MANAGER (GitHub Gist) ======
var SyncManager = {
  token: null,
  gistId: null,
  syncing: false,
  enabled: true,
  lastSync: null,
  pushTimer: null,

  // Simple hash for gist description
  hash: function(s) {
    var h = 0;
    for (var i = 0; i < s.length; i++) { h = ((h << 5) - h + s.charCodeAt(i)) | 0; }
    return Math.abs(h).toString(36);
  },

  api: function(path, opts) {
    var headers = { 'Accept': 'application/vnd.github+json' };
    if (this.token) headers['Authorization'] = 'Bearer ' + this.token;
    return fetch('https://api.github.com/' + path, {
      method: opts && opts.method || 'GET',
      headers: headers,
      body: opts && opts.body ? JSON.stringify(opts.body) : undefined
    }).then(function(r) {
      if (!r.ok) throw new Error('GitHub API ' + r.status);
      return r.json();
    });
  },

  login: function(username, token) {
    var self = this;
    self.token = token;
    return self.api('user').then(function(user) {
      // Verify username matches
      if (user.login.toLowerCase() !== username.toLowerCase()) {
        throw new Error('用户名与 Token 不匹配');
      }
      // Save credentials
      localStorage.setItem('_sync_token', token);
      localStorage.setItem('_sync_user', user.login);
      // Find or create gist
      return self.findOrCreateGist(user.login);
    }).then(function(gistId) {
      self.gistId = gistId;
      localStorage.setItem('_sync_gist', gistId);
      self.updateUI(true);
      // Pull data from gist
      return self.pull();
    });
  },

  findOrCreateGist: function(username) {
    var self = this;
    var desc = 'xi-gai-sync-' + self.hash(username);
    // Search for existing gist
    return self.api('gists?per_page=100').then(function(gists) {
      for (var i = 0; i < gists.length; i++) {
        if (gists[i].description === desc) return gists[i].id;
      }
      // Create new gist
      return self.api('gists', {
        method: 'POST',
        body: {
          description: desc,
          public: false,
          files: { 'sync.json': { content: '{}' } }
        }
      }).then(function(g) { return g.id; });
    });
  },

  logout: function() {
    this.token = null; this.gistId = null;
    localStorage.removeItem('_sync_token');
    localStorage.removeItem('_sync_user');
    localStorage.removeItem('_sync_gist');
    this.updateUI(false);
  },

  isloggedIn: function() {
    return !!localStorage.getItem('_sync_token');
  },

  autoLogin: function() {
    var token = localStorage.getItem('_sync_token');
    var user = localStorage.getItem('_sync_user');
    var gist = localStorage.getItem('_sync_gist');
    if (!token) return;
    this.token = token;
    this.gistId = gist;
    this.updateUI(true);
    // Verify token still works
    var self = this;
    this.api('user').then(function(u) {
      if (!gist) return self.findOrCreateGist(u.login).then(function(id) {
        self.gistId = id; localStorage.setItem('_sync_gist', id);
      });
    }).then(function() { return self.pull(); })
    .catch(function() { self.logout(); });
  },

  collectData: function() {
    var data = {};
    ['wrongBook', 'unsureBook', 'bookmarks', 'disputes'].forEach(function(k) {
      data[k] = loadJSON(k);
    });
    data._theme = localStorage.getItem('_theme') || 'light';
    data._users = JSON.parse(localStorage.getItem('_users') || '["默认用户"]');
    data._currentUser = currentUser;
    data._studySeconds = state.studySeconds;
    data._syncedAt = new Date().toISOString();
    return data;
  },

  applyData: function(data) {
    if (!data || !data._syncedAt) return;
    ['wrongBook', 'unsureBook', 'bookmarks', 'disputes'].forEach(function(k) {
      if (data[k]) saveJSON(k, data[k], true); // silent = no re-sync
    });
    if (data._theme) localStorage.setItem('_theme', data._theme);
    if (data._users) localStorage.setItem('_users', JSON.stringify(data._users));
    if (data._currentUser) {
      currentUser = data._currentUser;
      initUsers();
    }
    this.lastSync = data._syncedAt;
    updateBadge();
  },

  push: function() {
    if (!this.token || !this.gistId || !this.enabled) return;
    var self = this;
    clearTimeout(self.pushTimer);
    self.pushTimer = setTimeout(function() {
      self.syncing = true;
      self.updateUISyncing(true);
      var data = self.collectData();
      self.api('gists/' + self.gistId, {
        method: 'PATCH',
        body: { files: { 'sync.json': { content: JSON.stringify(data, null, 2) } } }
      }).then(function() {
        self.lastSync = new Date().toISOString();
        self.syncing = false;
        self.updateUISyncing(false);
      }).catch(function(e) {
        self.syncing = false;
        self.updateUISyncing(false);
        console.warn('Sync push failed:', e);
      });
    }, 2000);
  },

  pull: function() {
    if (!this.token || !this.gistId) return Promise.resolve();
    var self = this;
    self.syncing = true;
    self.updateUISyncing(true);
    return self.api('gists/' + self.gistId).then(function(gist) {
      var file = gist.files && gist.files['sync.json'];
      if (file && file.content) {
        try {
          var remote = JSON.parse(file.content);
          var local = self.collectData();
          // Merge: use newer timestamp per key
          if (remote._syncedAt && (!local._syncedAt || remote._syncedAt > local._syncedAt)) {
            self.applyData(remote);
          }
        } catch(e) {}
      }
      self.syncing = false;
      self.updateUISyncing(false);
    }).catch(function(e) {
      self.syncing = false;
      self.updateUISyncing(false);
    });
  },

  updateUI: function(loggedIn) {
    var dot = document.getElementById('sync-dot');
    var label = document.getElementById('sync-label');
    if (!dot || !label) return;
    if (loggedIn) {
      dot.className = 'dot online';
      label.textContent = localStorage.getItem('_sync_user') || '已登录';
    } else {
      dot.className = 'dot';
      label.textContent = '未登录';
    }
  },

  updateUISyncing: function(syncing) {
    var dot = document.getElementById('sync-dot');
    if (!dot) return;
    if (syncing) dot.className = 'dot syncing';
    else if (this.isloggedIn()) dot.className = 'dot online';
  }
};

function showSyncModal() {
  if (SyncManager.isloggedIn()) {
    // Show logout option
    if (confirm('当前登录: ' + localStorage.getItem('_sync_user') + '\n\n是否退出登录？（数据仍保留在本地）')) {
      SyncManager.logout();
    }
    return;
  }
  document.getElementById('login-modal').style.display = 'flex';
  document.getElementById('login-user').focus();
}

function doLogin() {
  var username = document.getElementById('login-user').value.trim();
  var token = document.getElementById('login-token').value.trim();
  var errEl = document.getElementById('login-error');
  errEl.style.display = 'none';

  if (!username || !token) {
    errEl.textContent = '请填写用户名和 Token'; errEl.style.display = 'block'; return;
  }
  if (!token.startsWith('ghp_') && !token.startsWith('github_pat_')) {
    errEl.textContent = 'Token 格式不对，应以 ghp_ 或 github_pat_ 开头'; errEl.style.display = 'block'; return;
  }

  errEl.textContent = '验证中...'; errEl.style.display = 'block'; errEl.style.color = 'var(--text2)';

  SyncManager.login(username, token).then(function() {
    document.getElementById('login-modal').style.display = 'none';
    errEl.style.display = 'none';
    // Re-render settings if on settings view
    if (state.view === 'settings') renderSettings();
  }).catch(function(e) {
    errEl.textContent = '登录失败: ' + (e.message || '未知错误');
    errEl.style.color = 'var(--danger)';
    errEl.style.display = 'block';
  });
}

// Auto-login on page load
SyncManager.autoLogin();
"""

# Insert sync_js before the init() call
html = html.replace('init();\n', 'init();\n' + sync_js + '\n')

# ============================================================
# 4. MODIFY saveJSON to trigger sync
# ============================================================
old_saveJSON = "function saveJSON(k, v) { localStorage.setItem(userKey(k), JSON.stringify(v)); }"
new_saveJSON = "function saveJSON(k, v, silent) { localStorage.setItem(userKey(k), JSON.stringify(v)); if (!silent && SyncManager) SyncManager.push(); }"
html = html.replace(old_saveJSON, new_saveJSON)

# ============================================================
# 5. UPDATE settings view to show sync status
# ============================================================
old_settings_about = """'<div class="settings-section"><h3>关于</h3><p style="font-size:13px;color:var(--text2)">选择题库 ' + (questionData?questionData.meta.totalQuestions:'?') + ' 题 · 大题 ' + (essayData?essayData.meta.totalQuestions:'?') + ' 题 · 卡片 ' + (flashcardData&&flashcardData.cards?flashcardData.cards.length:'?') + ' 张</p></div>';"""
new_settings_about = """'<div class="settings-section"><h3>跨设备同步</h3>' + (SyncManager.isloggedIn() ? '<p style="font-size:13px;color:var(--success)">✓ 已登录 ' + localStorage.getItem('_sync_user') + '</p><p style="font-size:12px;color:var(--text2);margin-top:4px">上次同步: ' + (SyncManager.lastSync ? new Date(SyncManager.lastSync).toLocaleString() : '等待中...') + '</p><button onclick="SyncManager.pull().then(function(){renderSettings()})" style="margin-top:8px">立即同步</button><button onclick="if(confirm(\'退出登录？\')){SyncManager.logout();renderSettings();}" style="color:var(--danger);margin-top:8px">退出登录</button>' : '<p style="font-size:13px;color:var(--text2)">登录 GitHub 账号即可跨设备同步学习数据</p><button onclick="showSyncModal()" style="margin-top:8px">登录以同步</button>') + '</div><div class="settings-section"><h3>关于</h3><p style="font-size:13px;color:var(--text2)">选择题库 ' + (questionData?questionData.meta.totalQuestions:'?') + ' 题 · 大题 ' + (essayData?essayData.meta.totalQuestions:'?') + ' 题 · 卡片 ' + (flashcardData&&flashcardData.cards?flashcardData.cards.length:'?') + ' 张 · 速览 ' + (slidesData&&slidesData.items?slidesData.items.length:'?') + ' 题</p></div>';"""
html = html.replace(old_settings_about, new_settings_about)

# Write back
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Sync feature added. File size: {len(html):,} bytes")

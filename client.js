/* dsh-arya-maildesk client face v2.13.0 (hand-written, zero build step).
 *
 * Two surfaces, one data path:
 *
 *   - `main` panel (key `mail`) — the mailbox itself: folder rail, message
 *     list with a real selection, reading pane, and organize actions
 *     (read state, move, delete). Reached from a `sidebar.panellist` icon.
 *   - `sidebar.footer.action` badge — the doorbell. New arrivals, and a
 *     click opens a mail in Outlook. It never grows into a second mailbox.
 *
 * Host calls go through POST /maildesk/api/call, an allowlisted route in the
 * host half. There is no send path here on purpose: the panel can organize a
 * mailbox, and only the agent's two-call confirmation can send from it.
 *
 * Loader format matches the modules node half: a __ModuleLoader__ bundle
 * whose factory receives `require` for kernel-provided modules.
 */
window.__ModuleLoader__.load({
  // Must be the package's own name, scope included: the module table keys every
  // bundle by package name, so an unscoped id here registers this module under a
  // name nothing imports and the web boot fails with
  // `client-modules: could not load "@aryacode/dsh-arya-maildesk"`.
  // test/bundle.test.mjs enforces this.
  id: '@aryacode/dsh-arya-maildesk',
  factory: (require) => {
    const React = require('react')
    const ReactDOM = require('react-dom')
    const h = React.createElement

    const STATE_URL = '/olk-notify/api/state'
    const CALL_URL = '/maildesk/api/call'
    const HEALTH_URL = '/maildesk/api/health'
    // First paint resolves the default Inbox by NAME, before the folder tree is
    // known; entryIds replace it as soon as a person picks a folder.
    const INBOX = 'Inbox'
    // Where the panel remembers its shape between opens.
    const UI_KEY = 'maildesk.ui.v1'
    const POLL_MS = 20000

    /** Lifecycle beacon → host debug file (diagnostics; silent on failure). */
    const beacon = (message) => {
      try {
        fetch('/olk-notify/api/client-log', {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ message }),
        }).catch(() => {})
      } catch { /* never break the UI over a beacon */ }
    }

    /** One host call. Resolves the bridge payload; rejects with its message. */
    async function call(action, args) {
      const res = await fetch(CALL_URL, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ action, args: args || {} }),
      })
      const data = await res.json().catch(() => null)
      if (data === null || data.ok !== true) {
        throw new Error((data && data.error) || ('HTTP ' + res.status))
      }
      const result = data.result || {}
      if (result.status === 'error') throw new Error(result.error || 'Outlook returned an error')
      return result
    }

    /** POST to one of the compose routes; resolves the parsed JSON either way. */
    async function api(path, body) {
      try {
        const res = await fetch(path, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify(body || {}),
        })
        const data = await res.json().catch(() => null)
        if (data === null || typeof data !== 'object') return { ok: false, error: 'HTTP ' + res.status }
        return data
      } catch (error) {
        return { ok: false, error: (error && error.message) || 'network error' }
      }
    }

    /** "Name <a@b>; c@d" ⇄ [{name,email}]. Tolerant: a bare address is enough. */
    function parsePeople(text) {
      return String(text || '').split(/[;,]/).map((x) => x.trim()).filter(Boolean).map((x) => {
        const m = /^(.*)<([^>]+)>\s*$/.exec(x)
        if (m) return { name: m[1].trim().replace(/^["']|["']$/g, ''), email: m[2].trim() }
        return { name: '', email: x }
      }).filter((p) => p.email.includes('@'))
    }
    function formatPeople(list) {
      return (Array.isArray(list) ? list : [])
        .filter((p) => p && p.email)
        .map((p) => (p.name ? p.name + ' <' + p.email + '>' : p.email))
        .join('; ')
    }
    /** The body box is plain text; each paragraph becomes a <p> on save. */
    function textToHtml(text) {
      const blocks = String(text || '').split(/\n{2,}/)
      return blocks.map((b) => '<p>' + b.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>') + '</p>').join('')
    }
    /** The inverse, for reopening a draft the host already holds: without it a
     *  reopened draft would look empty and the next save would erase the body. */
    function htmlToText(html) {
      return String(html || '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/(p|div|li|tr|h[1-6]|blockquote)>/gi, '\n\n')
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&')
        .replace(/\n{3,}/g, '\n\n')
        .trim()
    }

    function ensureStyles() {
      if (document.head === null) return null
      if (document.getElementById('dsh-arya-maildesk-styles') !== null) return null
      const style = document.createElement('style')
      style.id = 'dsh-arya-maildesk-styles'
      style.textContent = [
        /* --- sidebar row (badge) --- */
        '[data-slot="sidebar.footer.action"]{display:flex !important;flex-direction:column;width:100%;align-items:stretch;min-width:0}',
        '.md-trigger{box-sizing:border-box;cursor:pointer;flex:none;width:calc(100% + 4px);min-width:0;height:42px;color:var(--dsw-alias-label-primary);background:0 0;border:none;border-radius:12px;align-items:center;gap:8px;margin:4px -2px;padding:0 10px 0 8px;font-family:inherit;font-size:14px;line-height:22px;display:flex;transition:background .12s ease;position:relative}',
        '.md-trigger:hover,.md-trigger:focus-visible{background:var(--dsw-alias-interactive-bg-hover);outline:none}',
        '.md-trigger.md-rail{border-radius:50%;justify-content:center;gap:0;width:36px;height:36px;margin:8px 0 10px;padding:0}',
        '.md-label{white-space:nowrap;overflow:hidden;min-width:0;flex:1 1 auto;text-align:left}',
        '.md-chip{position:absolute;top:-4px;right:-4px;min-width:16px;height:16px;padding:0 4px;border-radius:8px;font-size:10px;color:#fff;background:var(--dsw-alias-label-danger,#e5484d);box-shadow:0 0 0 2px var(--dsw-specific-sidebar-fill,#fff);display:flex;align-items:center;justify-content:center;z-index:1}',
        '.md-chip[hidden]{display:none}',
        '.md-trigger.md-rail .md-chip{top:4px;right:4px;min-width:14px;height:14px;padding:0 3px;border-radius:7px;font-size:9px;box-shadow:none}',
        '.md-popover{position:fixed;z-index:60;min-width:260px;max-width:360px;padding:10px 12px;border-radius:10px;background-color:var(--dsh-bg-elevated,#fff);color:inherit;box-shadow:0 8px 24px rgba(0,0,0,.18);font-size:12.5px;line-height:1.6}',
        '.md-popover-title{font-weight:600;margin-bottom:6px}',
        '.md-popover-hint{opacity:.65;margin-top:6px}',
        '.md-mailrow-btn{display:block;width:100%;text-align:left;background:0 0;border:none;border-radius:6px;padding:4px 6px;margin:0 -6px;font:inherit;color:inherit;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;transition:background .12s ease}',
        '.md-mailrow-btn:hover:not(:disabled),.md-mailrow-btn:focus-visible:not(:disabled){background:var(--dsw-alias-interactive-bg-hover);outline:none}',
        '.md-mailrow-btn:disabled{cursor:default;opacity:.55}',
        '.md-toast{position:fixed;z-index:70;left:50%;bottom:28px;transform:translateX(-50%);max-width:min(70vw,520px);padding:8px 14px;border-radius:10px;background:var(--dsh-bg-elevated,#fff);color:var(--dsw-alias-label-primary,#222);box-shadow:0 8px 24px rgba(0,0,0,.18);font-size:12.5px;line-height:1.5;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
        '.md-compose{position:absolute;inset:0;z-index:6;display:flex;flex-direction:column;background:var(--dsh-bg,#fff)}',
        '.md-compose-h{display:flex;align-items:center;gap:8px;padding:10px 14px;border-bottom:1px solid var(--dsw-alias-border-subtle,#eee);font-weight:600;font-size:13.5px;flex:none}',
        '.md-compose-body{flex:1 1 auto;min-height:0;overflow:auto;padding:12px 14px;display:flex;flex-direction:column;gap:10px}',
        '.md-field{display:flex;gap:8px;align-items:flex-start}',
        '.md-field>label{flex:none;width:52px;padding-top:6px;font-size:12px;opacity:.7}',
        '.md-input,.md-textarea{flex:1 1 auto;min-width:0;box-sizing:border-box;padding:6px 9px;border-radius:7px;border:1px solid var(--dsw-alias-border-subtle,#dcdcdc);background:var(--dsh-bg,#fff);color:inherit;font:12.5px/1.5 inherit}',
        '.md-textarea{min-height:180px;resize:vertical;font-family:inherit}',
        '.md-compose-body textarea.md-textarea{flex:1 1 auto}',
        '.md-quote{margin-top:2px;padding:8px 10px;border-left:3px solid var(--dsw-alias-border-subtle,#ddd);opacity:.75;max-height:150px;overflow:auto;font-size:12px}',
        '.md-compose-actions{display:flex;gap:8px;align-items:center;padding:10px 14px;border-top:1px solid var(--dsw-alias-border-subtle,#eee);flex:none;flex-wrap:wrap}',
        '.md-link{background:none;border:0;color:var(--dsw-alias-label-accent,#2b6cb0);cursor:pointer;font:12px inherit;padding:0 4px}',
        '.md-modal{position:absolute;inset:0;z-index:8;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.34)}',
        '.md-modal-card{width:min(560px,92%);max-height:80%;overflow:auto;background:var(--dsh-bg,#fff);border-radius:12px;box-shadow:0 18px 48px rgba(0,0,0,.28);padding:16px 18px}',
        '.md-modal-card h3{margin:0 0 10px;font-size:14px}',
        '.md-confirm-row{display:flex;gap:8px;font-size:12.5px;padding:3px 0;border-bottom:1px dashed var(--dsw-alias-border-subtle,#eee)}',
        '.md-confirm-row b{flex:none;width:64px;opacity:.65;font-weight:600}',
        '.md-confirm-body{margin:10px 0 0;padding:10px;border-radius:8px;background:var(--dsw-alias-bg-subtle,#f7f7f7);white-space:pre-wrap;font-size:12.5px;max-height:220px;overflow:auto}',

        /* --- the draft card in chat --- */
        '.md-card{border:1px solid var(--dsw-alias-border-subtle,#e3e3e3);border-radius:10px;background:var(--dsh-bg-elevated,#fff);padding:10px 12px;margin:2px 0;max-width:640px;font:13px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif}',
        '.md-card-head{display:flex;align-items:center;gap:8px;min-width:0}',
        '.md-card-badge{font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;padding:2px 7px;border-radius:999px;background:#2f6fed14;color:#2f6fed;white-space:nowrap;flex:none}',
        '.md-card-title{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
        '.md-card-row{display:flex;gap:6px;margin-top:5px;font-size:12px;color:var(--dsw-alias-label-secondary,#5a6472)}',
        '.md-card-label{flex:none;min-width:16px;color:var(--dsw-alias-label-tertiary,#8a93a0)}',
        '.md-card-value{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
        '.md-card-body{margin-top:7px;color:var(--dsw-alias-label-primary,#1b1b1b);white-space:pre-wrap;max-height:140px;overflow:auto}',
        '.md-card-pending{color:var(--dsw-alias-label-secondary,#5a6472)}',
        '.md-card-done{opacity:.85}',
        '.md-card-error{margin-top:7px;color:#c0392b;font-size:12px}',
        '.md-card-confirm{margin-top:8px;padding:9px 10px;border-radius:8px;background:var(--dsw-alias-bg-subtle,#f7f7f7)}',
        '.md-card-confirm-head{font-weight:600;margin-bottom:5px}',
        '.md-card-confirm-row{font-size:12px;margin-top:3px}',
        '.md-card-confirm-note{margin-top:7px;font-size:11.5px;color:var(--dsw-alias-label-tertiary,#8a93a0)}',
        '.md-card-actions{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap}',

        /* --- the draft beside the chat (right-pane tab) --- */
        '.md-rt{display:flex;flex-direction:column;gap:2px;padding:12px 14px;font:13px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif;color:var(--dsw-alias-label-primary,#1b1b1b);height:100%;overflow:auto}',
        '.md-rt *{box-sizing:border-box}',
        '.md-rt-head{display:flex;align-items:baseline;gap:8px;margin-bottom:4px}',
        '.md-rt-title{font-size:13.5px;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
        '.md-rt-body{margin-top:9px;white-space:pre-wrap;border-top:1px solid var(--dsw-alias-border-subtle,#e6e6e6);padding-top:9px}',
        '.md-rt-note{margin:6px 0;color:var(--dsw-alias-label-secondary,#5a6472)}',
        '.md-rt-foot{margin-top:12px;border-top:1px solid var(--dsw-alias-border-subtle,#e6e6e6);padding-top:8px}',
        '.md-rt .md-card-actions{margin-top:12px}',
        '.md-rt .md-link{background:none;border:0;padding:0;color:var(--dsw-alias-label-tertiary,#8a93a0);font:inherit;font-size:12px;cursor:pointer;text-decoration:underline}',

        /* --- main panel --- */
        '.md{position:relative;display:flex;flex-direction:column;height:100%;min-height:0;font:13px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif;color:var(--dsw-alias-label-primary,#1b1b1b);background:var(--dsh-bg,#fff)}',
        '.md *{box-sizing:border-box}',
        '.md-top{display:flex;align-items:center;gap:10px;padding:10px 14px;border-bottom:1px solid var(--dsw-alias-border-subtle,#e6e6e6);flex:none}',
        '.md-title{font-size:14px;font-weight:600;letter-spacing:.1px}',
        '.md-pill{font-size:11px;padding:2px 8px;border-radius:999px;background:#8898ac18;color:var(--dsw-alias-label-secondary,#5a6472);white-space:nowrap}',
        '.md-pill.fixture{background:#eab30822;color:#8a6100}',
        '.md-pill.warn{background:#e5484d1f;color:#a3262b}',
        '.md-spacer{flex:1 1 auto}',
        '.md-btn{border:1px solid var(--dsw-alias-border-subtle,#d9d9d9);background:transparent;color:inherit;border-radius:8px;padding:5px 10px;font:inherit;font-size:12px;cursor:pointer;white-space:nowrap}',
        '.md-btn:hover:not(:disabled){background:var(--dsw-alias-interactive-bg-hover,#f2f2f2)}',
        '.md-btn:disabled{opacity:.4;cursor:default}',
        '.md-btn.primary{background:#3b6ef5;border-color:transparent;color:#fff}',
        '.md-btn.danger{color:#b42318;border-color:#e5484d55}',
        '.md-select{font:inherit;font-size:12px;padding:4px 6px;border-radius:8px;border:1px solid var(--dsw-alias-border-subtle,#d9d9d9);background:transparent;color:inherit}',
        '.md-body{flex:1 1 auto;display:flex;min-height:0}',
        '.md-rail{width:216px;flex:none;border-right:1px solid var(--dsw-alias-border-subtle,#e6e6e6);overflow:auto;padding:8px 6px}',
        '.md-rail-store{font-size:10.5px;text-transform:uppercase;letter-spacing:.6px;opacity:.55;padding:8px 8px 4px}',
        '.md-folder{display:flex;align-items:center;gap:6px;width:100%;text-align:left;background:0 0;border:none;border-radius:7px;padding:5px 8px;font:inherit;font-size:12.5px;color:inherit;cursor:pointer}',
        '.md-folder:hover{background:var(--dsw-alias-interactive-bg-hover,#f2f2f2)}',
        '.md-folder.on{background:#3b6ef514;color:#2350c9;font-weight:600}',
        '.md-folder-name{flex:1 1 auto;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
    // Lazy tree: a twisty that loads one level of children on demand. The rail
    // used to walk the whole mailbox (545 folders = 16 s); now it paints the
    // well-known folders and expands only what a person opens.
    '.md-twisty{flex:none;width:11px;text-align:center;font-size:9px;opacity:.55;transition:transform .12s;transform:rotate(0deg)}',
    '.md-twisty.open{transform:rotate(90deg)}',
    '.md-twisty.empty{opacity:0;cursor:default}',
        '.md-count{font-size:10.5px;opacity:.7;flex:none}',
        '.md-unread{font-size:10.5px;font-weight:700;color:#2350c9;flex:none}',
        '.md-list{flex:1 1 44%;min-width:320px;display:flex;flex-direction:column;border-right:1px solid var(--dsw-alias-border-subtle,#e6e6e6);min-height:0}',
        '.md-toolbar{display:flex;align-items:center;gap:6px;padding:8px 10px;border-bottom:1px solid var(--dsw-alias-border-subtle,#f0f0f0);flex:none;flex-wrap:wrap}',
        '.md-selinfo{font-size:11.5px;opacity:.7}',
        '.md-rows{flex:1 1 auto;overflow:auto;min-height:0}',
        '.md-row{display:grid;grid-template-columns:14px 1fr;gap:6px;padding:7px 10px;border-bottom:1px solid var(--dsw-alias-border-subtle,#f2f2f2);cursor:pointer;align-items:start}',
        '.md-row:hover{background:var(--dsw-alias-interactive-bg-hover,#f7f7f7)}',
        '.md-row.sel{background:#3b6ef516}',
        '.md-row.open{background:#3b6ef50e;box-shadow:inset 3px 0 0 #3b6ef5}',
        '.md-dot{width:7px;height:7px;border-radius:50%;background:#3b6ef5;margin-top:6px}',
        '.md-row.read .md-dot{background:transparent}',
        '.md-line1{display:flex;gap:8px;align-items:baseline;min-width:0}',
        '.md-from{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:52%}',
        '.md-row.read .md-from{font-weight:500;opacity:.85}',
        '.md-when{margin-left:auto;font-size:11px;opacity:.6;white-space:nowrap;flex:none}',
        '.md-subj{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12.5px}',
        '.md-prev{font-size:11.5px;opacity:.65;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
        '.md-clip{font-size:10.5px;opacity:.6;margin-left:4px}',
        '.md-read{flex:1 1 56%;min-width:340px;display:flex;flex-direction:column;min-height:0;overflow:hidden}',
        '.md-read-head{padding:12px 16px 10px;border-bottom:1px solid var(--dsw-alias-border-subtle,#f0f0f0);flex:none}',
        '.md-read-subj{font-size:15px;font-weight:600;line-height:1.35;margin-bottom:6px;word-break:break-word}',
        '.md-meta{font-size:11.5px;opacity:.72;line-height:1.6;word-break:break-word}',
        '.md-meta b{font-weight:600;opacity:.9}',
        '.md-att{display:inline-flex;align-items:center;gap:4px;font-size:11px;padding:2px 7px;border-radius:6px;background:#8898ac18;margin:4px 4px 0 0}',
        '.md-read-body{flex:1 1 auto;overflow:auto;padding:14px 16px 24px;white-space:pre-wrap;word-break:break-word;font-size:13px;line-height:1.6}',
        '.md-html{white-space:normal}',
        '.md-html p{margin:0 0 10px}',
        '.md-html table{border-collapse:collapse;margin:8px 0;font-size:12px}',
        '.md-html th,.md-html td{border:1px solid var(--dsw-alias-border-subtle,#dcdcdc);padding:4px 8px;text-align:left}',
        '.md-html blockquote{margin:8px 0;padding:2px 0 2px 10px;border-left:3px solid #8898ac55;opacity:.85}',
        '.md-html a{color:#3b6ef5;text-decoration:underline}',
        '.md-img-blocked{display:inline-block;font-size:11px;padding:2px 8px;border-radius:6px;background:#8898ac18;opacity:.8;border:1px dashed #8898ac55}',
        '.md-html-note{font-size:11px;opacity:.65;margin:0 0 8px}',
        '.md-empty{flex:1 1 auto;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;padding:28px;text-align:center;font-size:12.5px;opacity:.8}',
        '.md-empty-h{font-size:13.5px;font-weight:600;opacity:1}',
        '.md-err{margin:10px 14px;padding:10px 12px;border-radius:10px;background:#e5484d12;color:#a3262b;font-size:12px;line-height:1.55;white-space:pre-wrap;word-break:break-word}',
        '.md-skel{padding:7px 10px;border-bottom:1px solid #f2f2f2}',
        '.md-skel div{height:9px;border-radius:5px;background:#8898ac1f;margin:4px 0}',
        '.md-skel div:nth-child(1){width:34%}.md-skel div:nth-child(2){width:78%}.md-skel div:nth-child(3){width:60%;height:8px}',
        '.md-pager{display:flex;align-items:center;gap:8px;padding:8px 10px;border-top:1px solid var(--dsw-alias-border-subtle,#f0f0f0);font-size:11.5px;opacity:.8;flex:none}',
        '.md-kbd{font-size:10.5px;padding:1px 5px;border-radius:4px;border:1px solid var(--dsw-alias-border-subtle,#d9d9d9);background:#8898ac10}',
        '.md-ask{border-top:1px solid var(--dsw-alias-border-subtle,#e6e6e6);background:#fafbfc;display:flex;flex-direction:column;max-height:46%;flex:none}',
    '.md-ask-head{display:flex;align-items:center;gap:8px;padding:6px 10px;font-size:12px}',
    '.md-ask-sub{opacity:.6;font-size:11px}',
    '.md-ask-chips{display:flex;gap:6px;padding:0 10px 6px;flex-wrap:wrap}',
    '.md-ask-chip{font:inherit;font-size:11.5px;padding:3px 9px;border-radius:12px;border:1px solid #d8dde5;background:#fff;cursor:pointer;color:inherit}',
    '.md-ask-chip:hover:not(:disabled){background:#f2f5fb}',
    '.md-ask-chip:disabled{opacity:.55;cursor:default}',
    '.md-ask-log{overflow:auto;padding:6px 10px;display:flex;flex-direction:column;gap:6px;min-height:60px}',
    '.md-ask-msg{padding:6px 9px;border-radius:8px;font-size:12.5px;white-space:pre-wrap;max-width:86%;line-height:1.4}',
    '.md-ask-msg.user{align-self:flex-end;background:#3b6ef514;color:#2350c9}',
    '.md-ask-msg.assistant{align-self:flex-start;background:#fff;border:1px solid #e6e6e6}',
    '.md-ask-msg.tool{align-self:flex-start;background:transparent;border:1px dashed #d8dde5;color:#555;font-size:11.5px;padding:2px 7px}',
    '.md-move-list{max-height:300px;overflow:auto;border:1px solid #e6e6e6;border-radius:8px;margin:8px 0;background:#fff}',
    '.md-move-row{display:block;width:100%;text-align:left;padding:6px 9px;font-size:12px;background:none;border:0;border-bottom:1px solid #f2f2f2;cursor:pointer;color:inherit;font:inherit}',
    '.md-move-row:hover{background:#f4f6fa}',
    '.md-search{width:170px;flex:none;font:inherit;font-size:12px;padding:4px 8px;border:1px solid #d8dde5;border-radius:7px;background:#fff}',
    '.md-confirm-content{max-height:120px;overflow:auto;background:#fafbfc;border:1px solid #e6e6e6;border-radius:8px;padding:8px;font-size:12px;line-height:1.45;color:#333}',
    '.md-ask-empty{opacity:.65;font-size:12px;line-height:1.45}',
    '.md-ask-err{color:#e5484d;font-size:12px}',
    '.md-ask-compose{display:flex;gap:8px;padding:8px 10px;align-items:flex-end}',
    '.md-ask-input{flex:1 1 auto;font:inherit;font-size:12.5px;padding:6px 8px;border:1px solid #d8dde5;border-radius:8px;resize:vertical;background:#fff}',
    '@media (max-width:900px){.md-rail{display:none}.md-list{flex-basis:46%}.md-read{min-width:0}}',
      ].join('\n')
      document.head.appendChild(style)
      return style
    }

    /* ---------------------------------------------------------------- icons */

    function MailIcon({ size }) {
      return h('svg', {
        width: size, height: size, viewBox: '0 0 16 16',
        fill: 'none', stroke: 'currentColor', 'stroke-width': 1.3,
        'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': true,
      },
        h('rect', { x: 2, y: 3.5, width: 12, height: 9.5, rx: 1.5 }),
        h('path', { d: 'm2.5 4.5 5.5 4.7 5.5-4.7' }))
    }

    function FolderIcon({ size }) {
      return h('svg', {
        width: size, height: size, viewBox: '0 0 16 16',
        fill: 'none', stroke: 'currentColor', 'stroke-width': 1.2,
        'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': true,
      },
        h('path', { d: 'M2 4.5A1.5 1.5 0 0 1 3.5 3h2.2l1.3 1.6h5.5A1.5 1.5 0 0 1 14 6.1v5.4A1.5 1.5 0 0 1 12.5 13h-9A1.5 1.5 0 0 1 2 11.5z' }))
    }

    /* -------------------------------------------------------------- helpers */

    const shortWhen = (value) => {
      if (!value) return ''
      // "yyyy-MM-dd HH:mm" → today "09:41", otherwise "Mar 4"
      const [date, time] = String(value).split(' ')
      const today = new Date()
      const same = date === [today.getFullYear(), String(today.getMonth() + 1).padStart(2, '0'), String(today.getDate()).padStart(2, '0')].join('-')
      if (same && time) return time
      const [, m, d] = String(date).split('-')
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
      return months[Number(m) - 1] + ' ' + String(Number(d))
    }

    function toastHost(text, isError) {
      const div = document.createElement('div')
      div.className = 'md-toast'
      div.setAttribute('role', 'status')
      if (isError) div.style.color = 'var(--dsw-alias-label-danger,#e5484d)'
      div.textContent = text
      document.body.appendChild(div)
      setTimeout(() => { if (div.parentNode !== null) div.parentNode.removeChild(div) }, 2800)
    }

    /* ------------------------------------------------- reading pane sanitizer */

    // Mail HTML is hostile input: it comes from whoever wrote the message, and
    // it is rendered inside the harness frame. Everything below is an allowlist:
    // an element or attribute not named here does not survive.
    const DROP_WITH_CONTENT = new Set([
      'script', 'style', 'iframe', 'frame', 'frameset', 'object', 'embed', 'applet',
      'link', 'meta', 'base', 'title', 'head', 'form', 'input', 'button', 'select',
      'textarea', 'option', 'optgroup', 'label', 'fieldset', 'legend', 'svg', 'math',
      'video', 'audio', 'source', 'track', 'canvas', 'template', 'noscript', 'marquee',
    ])

    // Unknown tags are unwrapped (their text survives); these are kept.
    const KEEP_TAGS = new Set([
      'a', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'sub', 'sup', 'small', 'big',
      'p', 'div', 'span', 'br', 'hr', 'blockquote', 'pre', 'code', 'font', 'center',
      'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'dl', 'dt', 'dd',
      'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'caption', 'colgroup', 'col',
      'img', 'address', 'cite', 'q',
    ])

    const ATTRS_BY_TAG = {
      a: ['href', 'title'],
      img: ['alt', 'width', 'height'],
      td: ['colspan', 'rowspan', 'align', 'valign', 'width'],
      th: ['colspan', 'rowspan', 'align', 'valign', 'width'],
      table: ['border', 'cellpadding', 'cellspacing', 'width', 'align'],
      col: ['span', 'width'],
      colgroup: ['span', 'width'],
      font: ['color', 'face', 'size'],
      p: ['align'], div: ['align'], h1: ['align'], h2: ['align'], h3: ['align'],
    }

    // A style attribute is allowed, but only these declarations survive.
    const STYLE_PROPS = new Set([
      'color', 'background-color', 'font-weight', 'font-style', 'font-size', 'font-family',
      'text-decoration', 'text-align', 'text-transform', 'line-height', 'letter-spacing',
      'padding', 'padding-top', 'padding-right', 'padding-bottom', 'padding-left',
      'margin', 'margin-top', 'margin-right', 'margin-bottom', 'margin-left',
      'border', 'border-top', 'border-right', 'border-bottom', 'border-left',
      'border-color', 'border-style', 'border-width', 'border-collapse', 'border-radius',
      'width', 'max-width', 'height', 'vertical-align', 'white-space', 'display', 'float',
    ])

    const SAFE_URL = /^(?:https?:|mailto:|tel:)/i

    // Proofpoint rewrites every inbound link and wraps it in an urldefense hop.
    // Unwrap it so the reader sees and copies the address the sender wrote.
    function unwrapProofpoint(href) {
      let url = String(href || '')
      try {
        if (/urldefense\.proofpoint\.com\/v2\/url/i.test(url)) {
          const u = new URL(url).searchParams.get('u')
          if (u !== null && u !== '') url = decodeURIComponent(u)
        } else if (/urldefense\.(?:com|proofpoint\.com)\/v3\//i.test(url)) {
          const m = /\/v3\/__(.+?)__(?:;|$)/i.exec(url)
          if (m !== null) url = decodeURIComponent(m[1].replace(/_/g, '%5F'))
        }
      } catch { /* a malformed link is left exactly as it came */ }
      return url
    }

    function sanitizeStyle(value) {
      const kept = []
      for (const part of String(value).split(';')) {
        const i = part.indexOf(':')
        if (i < 0) continue
        const name = part.slice(0, i).trim().toLowerCase()
        const val = part.slice(i + 1).trim()
        if (!STYLE_PROPS.has(name)) continue
        if (val === '' || /url\s*\(|expression\s*\(|javascript:|\bimport\b|behavior\s*:/i.test(val)) continue
        kept.push(name + ': ' + val)
      }
      return kept.join('; ')
    }

    /**
     * Returns an HTML string that is safe to hand to the DOM: unknown elements
     * unwrapped, script-bearing elements dropped with their contents, every
     * attribute re-checked, remote images replaced by a placeholder so no
     * tracker is fetched, and Proofpoint hops unwrapped.
     */
    function sanitizeEmailHtml(html) {
      const doc = new DOMParser().parseFromString(String(html), 'text/html')
      const body = doc.body
      if (body === null) return ''

      const walk = (node) => {
        for (const child of Array.from(node.childNodes)) {
          if (child.nodeType === 3) continue
          if (child.nodeType !== 1) { child.remove(); continue }
          const tag = child.tagName.toLowerCase()
          if (DROP_WITH_CONTENT.has(tag)) { child.remove(); continue }

          if (!KEEP_TAGS.has(tag)) {
            // Unknown wrapper: keep the text, lose the element.
            walk(child)
            while (child.firstChild !== null) node.insertBefore(child.firstChild, child)
            child.remove()
            continue
          }

          walk(child)

          if (tag === 'img') {
            const alt = child.getAttribute('alt') || 'image'
            const chip = doc.createElement('span')
            chip.className = 'md-img-blocked'
            chip.textContent = '🖼 ' + alt + ' (remote image blocked)'
            child.replaceWith(chip)
            continue
          }

          const allowed = new Set(ATTRS_BY_TAG[tag] || [])
          for (const attr of Array.from(child.attributes)) {
            const name = attr.name.toLowerCase()
            if (name === 'style') {
              const safe = sanitizeStyle(attr.value)
              if (safe === '') child.removeAttribute(attr.name)
              else child.setAttribute('style', safe)
              continue
            }
            if (!allowed.has(name)) { child.removeAttribute(attr.name); continue }
            if (name === 'href') {
              const url = unwrapProofpoint(attr.value)
              if (!SAFE_URL.test(url)) { child.removeAttribute(attr.name); continue }
              child.setAttribute('href', url)
              child.setAttribute('rel', 'noopener noreferrer')
              child.setAttribute('target', '_blank')
              child.setAttribute('title', url)
            }
          }
        }
      }

      walk(body)
      return body.innerHTML
    }

    /* ---------------------------------------------------------- badge surface */

    function NewMailRow(props) {
      const wide = props.wide !== false
      const [pending, setPending] = React.useState([])
      const [unread, setUnread] = React.useState(null)
      const [open, setOpen] = React.useState(false)
      const [busy, setBusy] = React.useState(false)
      const btnRef = React.useRef(null)

      React.useEffect(() => {
        let live = true
        const load = async () => {
          try {
            const res = await fetch(STATE_URL)
            const data = await res.json()
            if (!live) return
            setPending(Array.isArray(data.pending) ? data.pending : [])
            setUnread(typeof data.unread === 'number' ? data.unread : null)
          } catch { /* host route not up yet */ }
        }
        void load()
        const timer = setInterval(() => { void load() }, POLL_MS)
        return () => { live = false; clearInterval(timer) }
      }, [])

      React.useEffect(() => {
        if (!open) return undefined
        const close = (ev) => {
          if (ev.type === 'keydown' && ev.key !== 'Escape') return
          if (ev.type === 'mousedown' && btnRef.current !== null && btnRef.current.contains(ev.target)) return
          setOpen(false)
        }
        document.addEventListener('mousedown', close)
        document.addEventListener('keydown', close)
        return () => {
          document.removeEventListener('mousedown', close)
          document.removeEventListener('keydown', close)
        }
      }, [open])

      const openMail = async (entryId) => {
        setBusy(true)
        try {
          const res = await fetch('/olk-notify/open', {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ entryId }),
          })
          const data = await res.json().catch(() => null)
          if (data && data.ok) toastHost('📨 Opened in Outlook: ' + (data.subject || '(no subject)'))
          else toastHost('⚠️ Could not open: ' + ((data && data.error) || ('HTTP ' + res.status)), true)
        } catch (error) {
          toastHost('⚠️ Could not open: ' + (error && error.message ? error.message : 'network error'), true)
        } finally {
          setBusy(false)
        }
      }

      const n = pending.length
      const lines = pending.slice(0, 8).map((m) => '[' + (m.sender || '?') + '] ' + (m.subject || '(no subject)'))
      if (typeof unread === 'number' && unread > 0) lines.unshift(unread + ' unread in Inbox')

      let popover = null
      if (open) {
        let left = 60
        let top = 60
        if (btnRef.current !== null) {
          const r = btnRef.current.getBoundingClientRect()
          left = r.right + 8
          top = Math.max(8, r.top - 8)
        }
        popover = ReactDOM.createPortal(
          h('div', {
            className: 'md-popover',
            role: 'dialog',
            'aria-label': 'New mail',
            style: { left: left + 'px', top: top + 'px' },
            onClick: (ev) => ev.stopPropagation(),
          },
            h('div', { className: 'md-popover-title', role: 'heading', 'aria-level': 2 }, '📬 New mail (' + n + ')'),
            pending.slice(0, 10).map((m, i) => h('button', {
              key: i,
              type: 'button',
              className: 'md-mailrow-btn',
              disabled: busy,
              onClick: () => { void openMail(m.entryId) },
              title: 'Click to open in Outlook and mark as read',
            },
              '[' + (m.sender || '?') + '] ' + (m.subject || '(no subject)') + ' · ' + (m.received || ''))),
            h('div', { className: 'md-popover-hint' }, 'Open Mail in the sidebar panel to read and organize; ask in chat to triage.')),
          document.body)
      }

      return h(React.Fragment, null,
        h('button', {
          ref: btnRef,
          type: 'button',
          className: 'md-trigger' + (wide ? '' : ' md-rail'),
          'data-dsh-plugin': 'dsh-arya-maildesk',
          'data-dsh-part': 'sidebar-entry',
          'aria-haspopup': 'dialog',
          'aria-label': n > 0 ? 'Mail, ' + n + ' new' : 'Mail',
          title: lines.length > 0 ? lines.join('\n') : 'Mail (nothing new)',
          onClick: () => setOpen((v) => !v),
        },
          h(MailIcon, { size: wide ? 16 : 18 }),
          wide ? h('span', { className: 'md-label' }, 'Mail') : null,
          h('span', { className: 'md-chip', hidden: n <= 0 }, String(n))),
        popover)
    }

    /* ----------------------------------------------------------- mail panel */

    function Panel() {
      const [folders, setFolders] = React.useState(null)
      // Lazy folder tree: children are fetched one folder at a time, so the
      // rail never waits for a whole-mailbox walk (545 folders = 16 s). The
      // tree, the expanded set and the folder you were reading are mirrored to
      // localStorage, so reopening Mail lands where you left it without a
      // single bridge call.
      const savedUi = (() => {
        try {
          const raw = localStorage.getItem(UI_KEY)
          const parsed = raw === null ? null : JSON.parse(raw)
          return parsed !== null && typeof parsed === 'object' ? parsed : {}
        } catch { return {} }
      })()
      const [kids, setKids] = React.useState(() => (savedUi.kids !== null && typeof savedUi.kids === 'object' ? savedUi.kids : {}))
      const [openFolders, setOpenFolders] = React.useState(() => (savedUi.openFolders !== null && typeof savedUi.openFolders === 'object' ? savedUi.openFolders : {}))
      const [folderId, setFolderId] = React.useState(() => (typeof savedUi.folderId === 'string' && savedUi.folderId !== '' ? savedUi.folderId : null))
      const [page, setPage] = React.useState({ total: 0, offset: 0, returned: 0, folderPath: '', mails: [], unreadInFolder: -1 })
      const [listState, setListState] = React.useState('loading')
      const [sel, setSel] = React.useState([])
      const [openId, setOpenId] = React.useState(null)
      const [openMail, setOpenMail] = React.useState(null)
      const [readState, setReadState] = React.useState('idle')
      const [error, setError] = React.useState('')
      const [busy, setBusy] = React.useState('')
      const [toast, setToast] = React.useState(null)
      // ---- the agent inside the panel -----------------------------------
      // The client bundle reaches the same session services the chat UI uses,
      // so Mail can run a real agent turn with the maildesk_* tools instead of
      // sending the person back to the workspace chat.
      const [askOpen, setAskOpen] = React.useState(false)
      const [askText, setAskText] = React.useState('')
      const [askLog, setAskLog] = React.useState([])
      const [askBusy, setAskBusy] = React.useState(false)
      const [askError, setAskError] = React.useState('')
      const askEndRef = React.useRef(null)
      const [unreadOnly, setUnreadOnly] = React.useState(false)
      const [sort, setSort] = React.useState('received')
      const [limit, setLimit] = React.useState(100)
      const [provider, setProvider] = React.useState('com')
      const [sample, setSample] = React.useState(false)
      // Move targets are their own thing now: the rail is lazy, so the picker
      // can load the whole tree once (cached host-side) instead of offering
      // only the folders a person happened to expand.
      const [moveOpen, setMoveOpen] = React.useState(false)
      const [moveFilter, setMoveFilter] = React.useState('')
      const [allFolders, setAllFolders] = React.useState(null)
      const [moveBusy, setMoveBusy] = React.useState(false)
      // Search is a mode over the same list: results replace the folder page
      // until cleared, and selection/actions work on them unchanged.
      const [searchText, setSearchText] = React.useState('')
      const [submittedSearch, setSubmittedSearch] = React.useState('')
      const [searchRows, setSearchRows] = React.useState(null)
      const [searchBusy, setSearchBusy] = React.useState(false)
      // The composer is a draft the HOST owns (see index.js compose routes):
      // this component only edits it and asks for a send. Nothing here creates
      // an Outlook item, so re-opening a reply cannot stack a second draft.
      const [compose, setCompose] = React.useState(null)
      const anchor = React.useRef(null)
      const rowsRef = React.useRef(null)

      // The chat draft card can ask this panel to open its draft. The request
      // may have arrived while the panel was unmounted, so drain it on mount.
      React.useEffect(() => {
        const unsubscribe = draftBus.subscribe((id) => { void openComposeById(id) })
        if (pendingMailDraft !== null) {
          const id = pendingMailDraft
          pendingMailDraft = null
          void openComposeById(id)
        }
        return unsubscribe
      }, [])

      const announce = (text, isError) => {
        setToast({ text, isError: isError === true })
        setTimeout(() => setToast(null), 2800)
      }

      // Load the mailbox: provider label, then the first page of mail, then the
      // folder tree. The order matters — the tree is the slowest call in the
      // plugin (one COM read per folder: 545 folders ≈ 16 s on a real Exchange
      // store), so the list must not wait for it. A folder NAME resolves inside
      // the bridge, so "Inbox" is enough to start.
      const boot = async () => {
        try {
          const res = await fetch(HEALTH_URL)
          const data = await res.json()
          if (data && data.ok === true) {
            setProvider(data.provider === 'fixture' ? 'fixture' : 'com')
            setSample(data.sample === true)
          }
        } catch { /* keep the com default */ }
        setFolderId((prev) => (prev === null || prev === '' ? INBOX : prev))
        // Let the list effect dispatch its request before the folder walk takes
        // the bridge queue (the bridge handles one call at a time).
        await new Promise((resolve) => setTimeout(resolve, 60))
        try {
          // Only the well-known folders of each store (~1.4 s). Subfolders are
          // fetched when a person expands one; the whole-tree walk was 16 s.
          const result = await call('list_folders', { roots: true })
          const list = Array.isArray(result.folders) ? result.folders : []
          setFolders(list)
          if (list.length === 0) setListState((s) => (s === 'loading' ? 'empty' : s))
        } catch (error) {
          // A failed folder walk must not blank a list that is already useful.
          setFolders([])
          announce('Folders could not be loaded: ' + error.message, true)
        }
      }

      // Flip the host between the sample mailbox and live COM, then reload.
      const switchSampleMail = async (on) => {
        setBusy('sample')
        try {
          const res = await fetch('/maildesk/api/fixture', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ on }),
          })
          const data = await res.json()
          if (!data || data.ok !== true) throw new Error((data && data.error) || 'the provider did not switch')
          setError('')
          setFolderId(null)
          setListState('loading')
          await boot()
          announce(on ? 'Showing sample mail — nothing here is real' : 'Back to live mail', false)
        } catch (error) {
          announce(error.message, true)
        } finally {
          setBusy('')
        }
      }

      React.useEffect(() => { void boot() }, [])

      React.useEffect(() => {
        if (folderId === null) return
        let live = true
        setListState('loading')
        setError('')
        void (async () => {
          try {
            const result = await call('list_mail', {
              folder: folderId,
              offset: 0,
              limit,
              sort,
              descending: true,
              preview: true,
              unreadOnly,
            })
            if (!live) return
            setPage({
              total: result.total || 0,
              offset: result.offset || 0,
              returned: result.returned || 0,
              folderPath: result.folderPath || '',
              mails: Array.isArray(result.mails) ? result.mails : [],
              unreadInFolder: typeof result.unreadInFolder === 'number' ? result.unreadInFolder : -1,
            })
            setListState('ready')
          } catch (error) {
            if (!live) return
            setError(error.message)
            setListState('error')
          }
        })()
        return () => { live = false }
      }, [folderId, unreadOnly, sort, limit])

      const refresh = async () => {
        if (folderId === null) return
        setBusy('refresh')
        try {
          const result = await call('list_mail', { folder: folderId, offset: 0, limit, sort, descending: true, preview: true, unreadOnly })
          setPage({
            total: result.total || 0,
            offset: 0,
            returned: result.returned || 0,
            folderPath: result.folderPath || '',
            mails: Array.isArray(result.mails) ? result.mails : [],
            unreadInFolder: typeof result.unreadInFolder === 'number' ? result.unreadInFolder : -1,
          })
          setListState('ready')
          setError('')
          const stillThere = new Set((result.mails || []).map((m) => m.entryId))
          setSel((prev) => prev.filter((id) => stillThere.has(id)))
        } catch (error) {
          setError(error.message)
          setListState('error')
        } finally {
          setBusy('')
        }
      }

      const loadMore = async () => {
        const offset = page.offset + page.returned
        setBusy('more')
        try {
          const result = await call('list_mail', { folder: folderId, offset, limit, sort, descending: true, preview: true, unreadOnly })
          setPage((prev) => ({
            ...prev,
            offset,
            returned: prev.returned + (result.returned || 0),
            total: result.total || prev.total,
            mails: prev.mails.concat(Array.isArray(result.mails) ? result.mails : []),
          }))
        } catch (error) {
          announce(error.message, true)
        } finally {
          setBusy('')
        }
      }

      const showMail = async (entryId) => {
        setOpenId(entryId)
        setOpenMail(null)
        setReadState('loading')
        try {
          const result = await call('read_mail', { entryId, maxChars: 20000, format: 'html' })
          setOpenMail({ entryId, ...result })
          setReadState('ready')
          // Opening a mail reads it, the way the desktop client does. The
          // toolbar can put it back, and nothing else changes state on read.
          const row = page.mails.find((m) => m.entryId === entryId)
          if (row !== undefined && row.unread === true) {
            try {
              await call('mark_read', { entryId, read: true })
              setPage((prev) => ({
                ...prev,
                unreadInFolder: Math.max(-1, (prev.unreadInFolder > 0 ? prev.unreadInFolder : 1) - 1),
                mails: prev.mails.map((m) => (m.entryId === entryId ? { ...m, unread: false } : m)),
              }))
            } catch { /* the mail is still readable if the flag write fails */ }
          }
        } catch (error) {
          setReadState('error')
          setOpenMail({ entryId, error: error.message })
        }
      }

      const pickRow = (mail, ev) => {
        const list = page.mails.map((m) => m.entryId)
        if (ev.shiftKey && anchor.current !== null) {
          const a = list.indexOf(anchor.current)
          const b = list.indexOf(mail.entryId)
          if (a >= 0 && b >= 0) {
            const range = list.slice(Math.min(a, b), Math.max(a, b) + 1)
            setSel(ev.ctrlKey || ev.metaKey ? (prev) => Array.from(new Set(prev.concat(range))) : range)
            return
          }
        }
        if (ev.ctrlKey || ev.metaKey) {
          setSel((prev) => (prev.includes(mail.entryId) ? prev.filter((id) => id !== mail.entryId) : prev.concat([mail.entryId])))
          anchor.current = mail.entryId
          return
        }
        setSel([mail.entryId])
        anchor.current = mail.entryId
        void showMail(mail.entryId)
      }

      const targetIds = () => (sel.length > 0 ? sel : (openId !== null ? [openId] : []))

      const doMark = async (read) => {
        const ids = targetIds()
        if (ids.length === 0) return
        setBusy('mark')
        try {
          await call('mark_read', { entryIds: ids, read })
          setPage((prev) => ({
            ...prev,
            mails: prev.mails.map((m) => (ids.includes(m.entryId) ? { ...m, unread: !read } : m)),
          }))
          announce((read ? 'Marked read' : 'Marked unread') + ' · ' + ids.length + (ids.length === 1 ? ' mail' : ' mails'))
        } catch (error) {
          announce(error.message, true)
        } finally {
          setBusy('')
        }
      }

      const doDelete = async () => {
        const ids = targetIds()
        if (ids.length === 0) return
        const ok = window.confirm('Move ' + ids.length + (ids.length === 1 ? ' mail' : ' mails') + ' to Deleted Items?\n\nThis is a soft delete: Outlook keeps them in Deleted Items until it empties that folder.')
        if (!ok) return
        setBusy('delete')
        try {
          // expected is the count gate: the bridge aborts if the selection it
          // is told about differs from the ids it received.
          const result = await call('delete_mail', { entryIds: ids, expected: ids.length })
          const gone = new Set(result.deletedEntryIds || ids)
          setPage((prev) => ({
            ...prev,
            total: Math.max(0, prev.total - gone.size),
            mails: prev.mails.filter((m) => !gone.has(m.entryId)),
            unreadInFolder: prev.unreadInFolder > 0 ? Math.max(0, prev.unreadInFolder - gone.size) : prev.unreadInFolder,
          }))
          setSel([])
          if (openId !== null && gone.has(openId)) { setOpenId(null); setOpenMail(null); setReadState('idle') }
          announce('Moved ' + (result.deleted || gone.size) + ' to Deleted Items · undo there')
          if (Array.isArray(result.failed) && result.failed.length > 0) announce(result.failed.length + ' could not be deleted', true)
        } catch (error) {
          announce(error.message, true)
        } finally {
          setBusy('')
        }
      }

      const moveSelection = async (destId) => {
        const ids = targetIds()
        if (ids.length === 0 || !destId) return
        setBusy('move')
        try {
          const result = await call('move_mail', { entryIds: ids, destFolder: destId })
          if (Array.isArray(result.failed) && result.failed.length > 0) {
            announce(result.failed.length + ' could not be moved: ' + result.failed[0].error, true)
          } else {
            announce('Moved ' + (result.moved || ids.length) + ' to ' + (result.destPath || 'folder'))
          }
          setSel([])
          setMoveOpen(false)
          if (openId !== null && ids.includes(openId)) { setOpenId(null); setOpenMail(null); setReadState('idle') }
          await refresh()
        } catch (error) {
          announce(error.message, true)
        } finally {
          setBusy('')
        }
      }

      const doOpenInOutlook = async () => {
        const ids = targetIds()
        if (ids.length === 0) return
        setBusy('outlook')
        try {
          for (const id of ids.slice(0, 10)) await call('open_mail', { entryId: id })
          announce('Opened ' + Math.min(ids.length, 10) + ' in Outlook')
        } catch (error) {
          announce(error.message, true)
        } finally {
          setBusy('')
        }
      }

      /* ---- composer ------------------------------------------------- */
      const DRAFT_URL = '/maildesk/api/draft'
      const SEND_URL = '/maildesk/api/send'

      /** Open a draft the host already holds — the same draft the chat card
       *  shows, so editing here and there is one object, not two. */
      async function openComposeById(draftId) {
        setCompose({ status: 'loading' })
        const res = await api(DRAFT_URL, { op: 'get', draftId })
        if (res.ok !== true) {
          setCompose({ status: 'error', error: res.error || 'that draft is no longer open' })
          return
        }
        const d = res.draft
        setCompose({
          status: 'edit',
          draft: d,
          reused: true,
          toText: formatPeople(d.to),
          ccText: formatPeople(d.cc),
          bccText: formatPeople(d.bcc),
          subject: d.subject || '',
          bodyText: htmlToText(d.bodyHtml),
          showCc: (d.cc || []).length > 0,
          showBcc: (d.bcc || []).length > 0,
          busy: '',
          error: '',
          confirm: null,
        })
      }

      const openCompose = async (mode) => {
        const ids = targetIds()
        const inReplyTo = ids.length === 1 ? ids[0] : null
        if (mode !== 'new' && inReplyTo === null) return
        setCompose({ status: 'loading' })
        const res = await api(DRAFT_URL, { op: 'open', mode, inReplyTo: inReplyTo || undefined })
        if (res.ok !== true) {
          setCompose({ status: 'error', error: (res.error || 'could not open a draft') + (res.result && res.result.error ? ': ' + res.result.error : '') })
          return
        }
        const d = res.draft
        setCompose({
          status: 'edit',
          draft: d,
          reused: res.reused === true,
          toText: formatPeople(d.to),
          ccText: formatPeople(d.cc),
          bccText: formatPeople(d.bcc),
          subject: d.subject || '',
          bodyText: htmlToText(d.bodyHtml),
          showCc: (d.cc || []).length > 0,
          showBcc: false,
          busy: '',
          error: '',
          confirm: null,
        })
      }

      /** Save what is on screen into the host draft, then ask to send. The
       *  host answers `needs-confirmation` with the hash of exactly this
       *  content; only a second call carrying that hash sends anything. */
      const doSend = async () => {
        if (compose === null || compose.status !== 'edit') return
        const c = compose
        const patch = {
          to: parsePeople(c.toText),
          cc: parsePeople(c.ccText),
          bcc: parsePeople(c.bccText),
          subject: c.subject,
          bodyHtml: textToHtml(c.bodyText),
        }
        setCompose({ ...c, busy: 'save', error: '' })
        const saved = await api(DRAFT_URL, { op: 'save', draftId: c.draft.id, patch })
        if (saved.ok !== true) {
          setCompose({ ...c, busy: '', error: saved.error || 'could not save the draft' })
          return
        }
        const first = await api(SEND_URL, { draftId: c.draft.id })
        const result = first.result || {}
        if (first.ok === true) {
          setCompose(null)
          announce('Sent · ' + (result.subject || saved.draft.subject || ''))
          void refresh()
          return
        }
        if (result.status === 'needs-confirmation') {
          setCompose({ ...c, draft: saved.draft, busy: '', confirm: { hash: result.hash, preview: result.preview } })
          return
        }
        setCompose({ ...c, draft: saved.draft, busy: '', error: result.error || first.error || 'send failed' })
      }

      const confirmSend = async () => {
        const c = compose
        if (c === null || c.confirm === null) return
        setCompose({ ...c, busy: 'send' })
        const res = await api(SEND_URL, { draftId: c.draft.id, confirmHash: c.confirm.hash })
        const result = res.result || {}
        if (res.ok === true) {
          setCompose(null)
          announce('Sent · ' + (result.subject || c.draft.subject || ''))
          void refresh()
          return
        }
        setCompose({
          ...c,
          busy: '',
          confirm: null,
          error: result.status === 'needs-confirmation'
            ? 'The draft changed while you were confirming — review it and send again.'
            : (result.error || res.error || 'send failed'),
        })
      }

      const discardDraft = async () => {
        const c = compose
        if (c === null || c.draft === undefined) { setCompose(null); return }
        setCompose({ ...c, busy: 'discard' })
        await api(DRAFT_URL, { op: 'discard', draftId: c.draft.id })
        setCompose(null)
        announce('Draft discarded')
      }

      const onKeyDown = (ev) => {
        if (compose !== null) { if (ev.key === 'Escape') { ev.preventDefault(); setCompose(null) } return }
        if (ev.key === 'Delete' && !ev.repeat) { ev.preventDefault(); void doDelete(); return }
        if (ev.key === 'u' || ev.key === 'U') { void doMark(true); return }
        if (ev.key === 'R' && ev.ctrlKey) { ev.preventDefault(); void refresh(); return }
        if (ev.key === 'Escape') { setSel([]); return }
        if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
          const list = page.mails.map((m) => m.entryId)
          if (list.length === 0) return
          ev.preventDefault()
          const current = openId === null ? -1 : list.indexOf(openId)
          const next = ev.key === 'ArrowDown' ? Math.min(list.length - 1, current + 1) : Math.max(0, current - 1)
          const mail = page.mails[next]
          if (mail !== undefined) { setSel([mail.entryId]); void showMail(mail.entryId) }
        }
      }

      const selectionCount = targetIds().length
      const noOutlook = /class not registered|80040154|not found|not installed|cannot create/i.test(error)

      /* ---- in-panel assistant ----------------------------------------- */
      // Ask posts to the plugin's own host route, which runs a real harness
      // agent turn in a session of its own and hands back the transcript. This
      // half never talks to a model directly and needs no session internals.
      const readAskLog = async () => {
        try {
          const res = await fetch('/maildesk/api/ask', { method: 'GET' })
          const data = await res.json()
          if (data === null || data.ok !== true) return
          const list = Array.isArray(data.messages) ? data.messages : []
          setAskLog(list.map((m) => ({ key: String(m.role) + String(m.seq), role: m.role, text: m.text })))
        } catch { /* the poll retries */ }
      }

      const askSend = async (textArg) => {
        const text = String(textArg !== undefined ? textArg : askText).trim()
        if (text === '' || askBusy) return
        setAskBusy(true)
        setAskError('')
        try {
          const res = await fetch('/maildesk/api/ask', {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({
              text,
              folderPath: page.folderPath || '',
              selected: openMail !== null && openMail.subject ? openMail.subject : '',
              selectedId: openMail !== null && openMail.entryId ? openMail.entryId : '',
            }),
          })
          const data = await res.json().catch(() => null)
          if (data === null || data.ok !== true) {
            throw new Error((data && data.error) || ('HTTP ' + res.status))
          }
          setAskText('')
          await readAskLog()
        } catch (error) {
          setAskError(error && error.message ? error.message : String(error))
        } finally {
          setAskBusy(false)
        }
      }

      // Poll while the pane is open: the agent answers asynchronously and the
      // host owns the session, so a 1.5 s read keeps the two in step.
      React.useEffect(() => {
        if (!askOpen) return undefined
        let live = true
        let timer = null
        const tick = async () => {
          if (!live) return
          await readAskLog()
          if (live) timer = setTimeout(() => void tick(), 1500)
        }
        void tick()
        return () => {
          live = false
          if (timer !== null) clearTimeout(timer)
        }
      }, [askOpen])

      React.useEffect(() => {
        if (askOpen && askEndRef.current !== null) askEndRef.current.scrollIntoView({ block: 'end' })
      }, [askOpen, askLog.length])

      // Mirror the rail's shape and the open folder after every change, so the
      // next open shows the same tree without waiting on the bridge.
      React.useEffect(() => {
        try {
          localStorage.setItem(UI_KEY, JSON.stringify({ folderId, kids, openFolders }))
        } catch { /* quota, or private mode: the panel still works */ }
      }, [folderId, kids, openFolders])


      // While folderId is still the "Inbox" NAME (first paint), no entryId can
      // match: highlight the folder the page actually came from instead.
      const isActiveFolder = (f) =>
        f.entryId === folderId ||
        (typeof folderId === 'string' && f.path === page.folderPath)

      // One expand = one small bridge call (children of a single folder, ~2 s
      // including process start) instead of the 16 s whole-tree walk.
      const toggleFolder = async (f) => {
        const id = f.entryId
        if (openFolders[id] === true) {
          setOpenFolders((o) => ({ ...o, [id]: false }))
          return
        }
        setOpenFolders((o) => ({ ...o, [id]: true }))
        if (Array.isArray(kids[id])) return
        try {
          const result = await call('list_folders', { parent: id })
          const depth = (f.depth || 0) + 1
          const list = (Array.isArray(result.folders) ? result.folders : []).map((c) => ({ ...c, depth }))
          setKids((k) => ({ ...k, [id]: list }))
        } catch (error) {
          announce(error.message, true)
          setOpenFolders((o) => ({ ...o, [id]: false }))
        }
      }

      // Everything the panel knows about, in tree order: the roots plus any
      // level a person has opened. The move picker offers exactly this.
      const flatFolders = (() => {
        const out = []
        const walk = (f) => {
          out.push(f)
          const k = kids[f.entryId]
          if (Array.isArray(k)) for (const c of k) walk(c)
        }
        for (const f of folders || []) walk(f)
        return out
      })()

      const rail = h('div', { className: 'md-rail', role: 'navigation', 'aria-label': 'Folders' },
        folders === null ? [0, 1, 2, 3].map((i) => h('div', { key: i, className: 'md-skel' }, h('div'), h('div'), h('div'))) : null,
        folders !== null ? (() => {
          const out = []
          const seenStore = new Set()
          const walk = (f) => {
            if (!seenStore.has(f.store)) {
              seenStore.add(f.store)
              out.push(h('div', { key: 'store-' + f.store, className: 'md-rail-store' }, f.store))
            }
            const id = f.entryId
            const unread = typeof f.unread === 'number' ? f.unread : 0
            const hasKids = (f.childCount || 0) > 0
            const expanded = openFolders[id] === true
            out.push(h('button', {
              key: id,
              type: 'button',
              className: 'md-folder' + (isActiveFolder(f) ? ' on' : ''),
              style: { paddingLeft: 8 + (f.depth || 0) * 12 + 'px' },
              title: f.path,
              onClick: () => { setFolderId(id); setSel([]); setOpenId(null); setOpenMail(null); setReadState('idle') },
            },
              h('span', {
                key: 'tw',
                className: 'md-twisty' + (hasKids ? '' : ' empty') + (expanded ? ' open' : ''),
                role: 'button',
                tabIndex: hasKids ? 0 : -1,
                'aria-label': (expanded ? 'Collapse ' : 'Expand ') + f.name,
                onClick: (ev) => { ev.stopPropagation(); if (hasKids) void toggleFolder(f) },
                onKeyDown: (ev) => {
                  if (hasKids && (ev.key === 'Enter' || ev.key === ' ')) {
                    ev.preventDefault()
                    ev.stopPropagation()
                    void toggleFolder(f)
                  }
                },
              }, hasKids ? '▸' : ''),
              h(FolderIcon, { size: 13 }),
              h('span', { className: 'md-folder-name' }, f.name),
              unread > 0 ? h('span', { className: 'md-unread' }, String(unread)) : null,
              h('span', { className: 'md-count' }, String(typeof f.itemCount === 'number' ? f.itemCount : ''))))
            if (expanded && Array.isArray(kids[id])) for (const c of kids[id]) walk(c)
          }
          for (const f of folders) walk(f)
          return out
        })() : null,
        h('div', { style: { padding: '10px 8px 2px', fontSize: '10.5px', opacity: .6 } },
          'Badge counts new arrivals only.'),
      )

      const toolbar = h('div', { className: 'md-toolbar' },
        h('button', { className: 'md-btn primary', type: 'button', disabled: busy !== '' || selectionCount === 0, onClick: () => void doMark(true) }, 'Mark read'),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || selectionCount === 0, onClick: () => void doMark(false) }, 'Mark unread'),
        h('button', { className: 'md-btn danger', type: 'button', disabled: busy !== '' || selectionCount === 0, onClick: () => void doDelete() }, 'Delete'),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || selectionCount === 0, title: 'Move the selection to a folder', onClick: () => { setMoveFilter(''); setMoveOpen(true) } }, 'Move to...'),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || selectionCount === 0, onClick: () => void doOpenInOutlook() }, 'Open in Outlook'),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || selectionCount !== 1, onClick: () => void openCompose('reply') }, 'Reply'),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || selectionCount !== 1, onClick: () => void openCompose('replyAll') }, 'Reply all'),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || selectionCount !== 1, onClick: () => void openCompose('forward') }, 'Forward'),
        h('span', { className: 'md-selinfo' }, selectionCount > 0 ? selectionCount + ' selected' : 'no selection'))

      // Outlook's own filter engine, reached the same way the tool is: subject
      // first, then sender, because a person typing "Dana" means either.
      const runSearch = async (text) => {
        const query = text.trim()
        if (query === '') {
          setSearchRows(null)
          setSubmittedSearch('')
          return
        }
        setSearchBusy(true)
        setSubmittedSearch(query)
        setSel([])
        setOpenId(null)
        setOpenMail(null)
        setReadState('idle')
        try {
          let result = await call('search_mail', { subjectContains: query, limit: 100, days: 365 })
          let mails = Array.isArray(result.mails) ? result.mails : []
          if (mails.length === 0) {
            result = await call('search_mail', { from: query, limit: 100, days: 365 })
            mails = Array.isArray(result.mails) ? result.mails : []
          }
          setSearchRows(mails)
        } catch (error) {
          setSearchRows([])
          announce(error.message, true)
        } finally {
          setSearchBusy(false)
        }
      }

      const clearSearch = () => {
        setSearchText('')
        setSubmittedSearch('')
        setSearchRows(null)
      }

      const listMails = searchRows !== null ? searchRows : page.mails

      const rows = (() => {
        if (listState === 'loading' && searchRows === null) {
          return [0, 1, 2, 3, 4, 5, 6].map((i) => h('div', { key: i, className: 'md-skel' }, h('div'), h('div'), h('div')))
        }
        if (listState === 'error' && searchRows === null) return null
        if (listMails.length === 0) {
          return h('div', { className: 'md-empty' },
            h('div', { className: 'md-empty-h' }, searchRows !== null
              ? 'Nothing matches "' + submittedSearch + '"'
              : (unreadOnly ? 'No unread mail here' : 'This folder is empty')),
            h('div', null, searchRows !== null ? 'Search covers this folder; try another word or clear the search.' : (page.folderPath || '')))
        }
        return listMails.map((m) => h('div', {
          key: m.entryId,
          className: 'md-row' + (m.unread ? '' : ' read') + (sel.includes(m.entryId) ? ' sel' : '') + (m.entryId === openId ? ' open' : ''),
          role: 'option',
          'aria-selected': sel.includes(m.entryId),
          tabIndex: 0,
          onClick: (ev) => pickRow(m, ev),
          onKeyDown: (ev) => { if (ev.key === 'Enter') pickRow(m, ev) },
        },
          h('div', { className: 'md-dot' }),
          h('div', { style: { minWidth: 0 } },
            h('div', { className: 'md-line1' },
              h('span', { className: 'md-from' }, m.sender || '(unknown sender)'),
              h('span', { className: 'md-when' }, shortWhen(m.received))),
            h('div', { className: 'md-subj' }, (m.hasAttach ? '📎 ' : '') + (m.subject || '(no subject)')),
            m.preview ? h('div', { className: 'md-prev' }, m.preview) : null)))
      })()

      const reader = (() => {
        if (readState === 'idle') {
          return h('div', { className: 'md-empty' },
            h('div', { className: 'md-empty-h' }, 'No mail selected'),
            h('div', null, 'Pick a message to read it here.'),
            h('div', null, h('span', { className: 'md-kbd' }, '↑'), ' ', h('span', { className: 'md-kbd' }, '↓'), ' move · ', h('span', { className: 'md-kbd' }, 'Del'), ' delete · ', h('span', { className: 'md-kbd' }, 'U'), ' unread'))
        }
        if (readState === 'loading') return h('div', { className: 'md-empty' }, 'Opening…')
        if (openMail !== null && openMail.error !== undefined) {
          return h('div', { className: 'md-empty' }, h('div', { className: 'md-empty-h' }, 'Could not read this mail'), h('div', null, openMail.error))
        }
        if (openMail === null) return null
        const attachments = Array.isArray(openMail.attachments) ? openMail.attachments.filter(Boolean) : []
        const safeHtml = typeof openMail.htmlBody === 'string' && openMail.htmlBody !== '' ? sanitizeEmailHtml(openMail.htmlBody) : null
        return h(React.Fragment, null,
          h('div', { className: 'md-read-head' },
            h('div', { className: 'md-read-subj' }, openMail.subject || '(no subject)'),
            h('div', { className: 'md-meta' },
              h('div', null, h('b', null, 'From: '), openMail.sender || ''),
              h('div', null, h('b', null, 'To: '), openMail.to || ''),
              openMail.cc ? h('div', null, h('b', null, 'Cc: '), openMail.cc) : null,
              h('div', null, h('b', null, 'Received: '), openMail.received || '')),
            attachments.length > 0 ? h('div', null, attachments.map((a, i) => h('span', { key: i, className: 'md-att' }, '📎 ' + a))) : null),
          h('div', { className: 'md-read-body' },
            openMail.htmlTruncated === true
              ? h('div', { className: 'md-html-note' }, 'The formatted body was truncated; text past the cut is missing.')
              : null,
            safeHtml !== null
              ? h('div', { className: 'md-html', dangerouslySetInnerHTML: { __html: safeHtml } })
              : h('div', { className: 'md-text' }, openMail.body || '(no text body)')))
      })()

      const composeView = (() => {
        if (compose === null) return null
        if (compose.status === 'loading') {
          return h('div', { className: 'md-compose' }, h('div', { className: 'md-compose-h' }, 'Opening draft…'))
        }
        if (compose.status === 'error') {
          return h('div', { className: 'md-compose' },
            h('div', { className: 'md-compose-h' }, 'Draft could not be opened'),
            h('div', { className: 'md-compose-body' }, h('div', { className: 'md-err', role: 'alert' }, compose.error)),
            h('div', { className: 'md-compose-actions' },
              h('button', { className: 'md-btn', type: 'button', onClick: () => setCompose(null) }, 'Close')))
        }
        const c = compose
        const title = c.draft.mode === 'replyAll' ? 'Reply all' : c.draft.mode === 'reply' ? 'Reply'
          : c.draft.mode === 'forward' ? 'Forward' : 'New message'
        const field = (label, value, onChange, placeholder) => h('div', { className: 'md-field' },
          h('label', null, label),
          h('input', {
            className: 'md-input', value,
            placeholder: placeholder === undefined ? 'Name <address>; another@example.com' : placeholder,
            onChange: (ev) => onChange(ev.target.value),
          }))
        return h('div', { className: 'md-compose', role: 'dialog', 'aria-label': title },
          h('div', { className: 'md-compose-h' }, title,
            c.reused ? h('span', { className: 'md-pill', title: 'This thread already had a draft; it was reused instead of creating another' }, 'draft kept') : null,
            h('span', { className: 'md-spacer' }),
            c.busy !== '' ? h('span', { className: 'md-selinfo' }, c.busy === 'send' ? 'Sending…' : c.busy === 'save' ? 'Saving…' : 'Discarding…') : null),
          h('div', { className: 'md-compose-body' },
            field('To', c.toText, (v) => setCompose({ ...c, toText: v })),
            c.showCc
              ? field('Cc', c.ccText, (v) => setCompose({ ...c, ccText: v }))
              : h('div', null,
                  h('button', { className: 'md-link', type: 'button', onClick: () => setCompose({ ...c, showCc: true }) }, 'Add Cc'),
                  c.showBcc ? null : h('button', { className: 'md-link', type: 'button', onClick: () => setCompose({ ...c, showBcc: true }) }, 'Add Bcc')),
            c.showBcc ? field('Bcc', c.bccText, (v) => setCompose({ ...c, bccText: v })) : null,
            field('Subject', c.subject, (v) => setCompose({ ...c, subject: v }), ''),
            h('div', { className: 'md-field' },
              h('label', null, 'Body'),
              h('textarea', {
                className: 'md-textarea',
                value: c.bodyText,
                placeholder: 'Write the reply. It is sent as HTML, with the quoted mail appended below it.',
                onChange: (ev) => setCompose({ ...c, bodyText: ev.target.value }),
              })),
            c.draft.quoteHtml
              ? h('div', { className: 'md-quote' },
                  h('div', { style: { opacity: .7, marginBottom: 4 } }, 'Quoted below in the sent message:'),
                  h('div', { dangerouslySetInnerHTML: { __html: sanitizeEmailHtml(c.draft.quoteHtml) } }))
              : null,
            c.error !== '' ? h('div', { className: 'md-err', role: 'alert' }, c.error) : null),
          h('div', { className: 'md-compose-actions' },
            h('button', { className: 'md-btn primary', type: 'button', disabled: c.busy !== '' || c.toText.trim() === '', onClick: () => void doSend() }, 'Send…'),
            h('button', { className: 'md-btn', type: 'button', disabled: c.busy !== '', onClick: () => void discardDraft() }, 'Discard'),
            h('span', { className: 'md-spacer' }),
            h('button', { className: 'md-btn', type: 'button', disabled: c.busy !== '', onClick: () => setCompose(null) }, 'Close'),
            h('span', { className: 'md-selinfo' }, 'Nothing leaves until you confirm the exact content.')),
          c.confirm !== null
            ? h('div', { className: 'md-modal' },
                h('div', { className: 'md-modal-card', role: 'alertdialog', 'aria-label': 'Confirm send' },
                  h('h3', null, 'Send this?'),
                  h('div', { className: 'md-confirm-row' }, h('b', null, 'To'), h('span', null, c.confirm.preview.to || '(none)')),
                  c.confirm.preview.cc ? h('div', { className: 'md-confirm-row' }, h('b', null, 'Cc'), h('span', null, c.confirm.preview.cc)) : null,
                  c.confirm.preview.bcc ? h('div', { className: 'md-confirm-row' }, h('b', null, 'Bcc'), h('span', null, c.confirm.preview.bcc)) : null,
                  h('div', { className: 'md-confirm-row' }, h('b', null, 'Subject'), h('span', null, c.confirm.preview.subject || '(no subject)')),
                  // The gate promises "what will leave"; the body is part of
                  // that, so show it instead of only naming the recipient.
                  h('div', { className: 'md-confirm-content' },
                    String(c.confirm.preview.htmlBody || '')
                      .replace(/<[^>]+>/g, ' ')
                      .replace(/&nbsp;/g, ' ')
                      .replace(/&amp;/g, '&')
                      .replace(/\s+/g, ' ')
                      .trim()
                    || '(empty body)'),
                  h('div', { className: 'md-confirm-body' }, 'The message is sent as HTML and includes the quoted mail below your text. Only this exact content is sent — editing it afterwards requires a new confirmation.'),
                  h('div', { className: 'md-compose-actions', style: { borderTop: 0, padding: '12px 0 0' } },
                    h('button', { className: 'md-btn primary', type: 'button', disabled: c.busy !== '', onClick: () => void confirmSend() }, c.busy === 'send' ? 'Sending…' : 'Send now'),
                    h('button', { className: 'md-btn', type: 'button', disabled: c.busy !== '', onClick: () => setCompose({ ...c, confirm: null }) }, 'Back'))))
            : null)
      })()

      // The agent lives where the mail is: same page, same mailbox, no tab
      // switching. It runs as a real harness session, so it keeps the full
      // maildesk_* tool surface and the same confirmation gates as the chat.
      const askDrawer = askOpen
        ? h('div', { className: 'md-ask' },
            h('div', { className: 'md-ask-head' },
              h('b', null, 'Mail assistant'),
              h('span', { className: 'md-ask-sub' }, 'same agent, same mail tools'),
              h('span', { className: 'md-spacer' }),
              h('button', { className: 'md-btn', type: 'button', onClick: () => setAskOpen(false) }, 'Close')),
            // With a mail open, the three things people actually ask are one
            // click away instead of typing the same sentence again.
            openMail !== null
              ? h('div', { className: 'md-ask-chips' },
                  ['Summarize this mail', 'What does it want from me?', 'Draft a short reply'].map((q) => h('button', {
                    key: q,
                    type: 'button',
                    className: 'md-ask-chip',
                    disabled: askBusy,
                    onClick: () => { setAskText(q); void askSend(q) },
                  }, q)))
              : null,
            h('div', { className: 'md-ask-log' },
              askLog.length === 0
                ? h('div', { className: 'md-ask-empty' },
                    'Ask about this mailbox — “summarize this thread”, “reply that Tuesday works”, “move this to Concur”. '
                    + 'The assistant is the same agent as the chat, with the maildesk tools; anything sent still needs your confirmation.')
                : askLog.map((m) => h('div', { key: m.key, className: 'md-ask-msg ' + m.role },
                    m.role === 'tool'
                      ? '⚙ ' + String(m.text)
                      // The transcript is plain text; drop the two markers that
                      // would otherwise show up literally in agent replies.
                      : String(m.text).replace(/\*\*(.+?)\*\*/g, '$1').replace(/^#{1,6}\s+/gm, ''))),
              askError !== '' ? h('div', { className: 'md-ask-err' }, askError) : null,
              h('div', { ref: askEndRef })),
            h('div', { className: 'md-ask-compose' },
              h('textarea', {
                className: 'md-ask-input',
                placeholder: 'Ask the mail assistant…  (Enter to send, Shift+Enter for a new line)',
                value: askText,
                rows: 2,
                onChange: (ev) => setAskText(ev.target.value),
                onKeyDown: (ev) => {
                  if (ev.key === 'Enter' && !ev.shiftKey) {
                    ev.preventDefault()
                    ev.stopPropagation()
                    void askSend()
                  }
                },
              }),
              h('button', {
                className: 'md-btn primary',
                type: 'button',
                disabled: askBusy || askText.trim() === '',
                onClick: () => void askSend(),
              }, askBusy ? 'Sending…' : 'Ask')))
        : null

      // The rail is lazy, so the move picker owns its own tree: the folders a
      // person has opened, plus a button that loads the whole mailbox once
      // (cached host-side, so it is slow at most once per five minutes).
      const loadAllFolders = async () => {
        if (allFolders !== null || moveBusy) return
        setMoveBusy(true)
        try {
          const result = await call('list_folders', { maxDepth: 6 })
          setAllFolders(Array.isArray(result.folders) ? result.folders : [])
        } catch (error) {
          announce(error.message, true)
        } finally {
          setMoveBusy(false)
        }
      }

      const moveQuery = moveFilter.trim().toLowerCase()
      const moveCandidates = (allFolders !== null ? allFolders : flatFolders).filter((f) =>
        moveQuery === '' || String(f.path || f.name || '').toLowerCase().includes(moveQuery))

      const moveDialog = moveOpen
        ? h('div', { className: 'md-modal' },
            h('div', { className: 'md-modal-card', role: 'dialog', 'aria-label': 'Move to folder' },
              h('h3', null, 'Move ' + selectionCount + ' to a folder'),
              h('input', {
                className: 'md-ask-input',
                placeholder: 'Filter folders — a name or a path',
                value: moveFilter,
                onChange: (ev) => setMoveFilter(ev.target.value),
              }),
              allFolders === null
                ? h('div', { style: { margin: '8px 0', display: 'flex', alignItems: 'center', gap: 8 } },
                    h('button', { className: 'md-btn', type: 'button', disabled: moveBusy, onClick: () => void loadAllFolders() },
                      moveBusy ? 'Loading every folder…' : 'Load every folder'),
                    h('span', { className: 'md-ask-sub' }, 'otherwise only folders you have opened are listed'))
                : null,
              h('div', { className: 'md-move-list' },
                moveCandidates.slice(0, 400).map((f) => h('button', {
                  key: f.entryId,
                  type: 'button',
                  className: 'md-move-row',
                  title: f.path,
                  onClick: () => void moveSelection(f.entryId),
                }, f.path || f.name))),
              moveCandidates.length === 0 ? h('div', { className: 'md-ask-empty' }, 'No folder matches that filter.') : null,
              h('div', { className: 'md-compose-actions', style: { borderTop: 0 } },
                h('button', { className: 'md-btn', type: 'button', onClick: () => setMoveOpen(false) }, 'Cancel'))))
        : null

      return h('div', { className: 'md', onKeyDown, tabIndex: -1 },
        h('div', { className: 'md-top' },
          h('span', { className: 'md-title' }, 'Mail'),
          provider === 'fixture'
            ? h('span', {
                className: 'md-pill fixture',
                title: sample
                  ? 'Sample mailbox — synthetic mail, so you can look at the interface on a machine without classic Outlook'
                  : 'MAILDESK_FIXTURE=1 — this is synthetic mail, not a mailbox',
              }, sample ? 'SAMPLE MAIL' : 'FIXTURE DATA')
            : h('span', { className: 'md-pill' }, 'Classic Outlook · COM'),
          sample
            ? h('button', { className: 'md-btn', type: 'button', disabled: busy !== '', onClick: () => void switchSampleMail(false) }, 'Back to live mail')
            : null,
          page.folderPath ? h('span', { className: 'md-pill' }, page.folderPath) : null,
          h('input', {
            className: 'md-search',
            type: 'search',
            placeholder: 'Search mail…',
            value: searchText,
            title: 'Search this folder by subject, then by sender',
            onChange: (ev) => setSearchText(ev.target.value),
            onKeyDown: (ev) => {
              if (ev.key === 'Enter') {
                ev.preventDefault()
                ev.stopPropagation()
                void runSearch(searchText)
              } else if (ev.key === 'Escape') {
                ev.stopPropagation()
                clearSearch()
              }
            },
          }),
          h('span', { className: 'md-spacer' }),
          h('label', { className: 'md-selinfo', style: { display: 'flex', alignItems: 'center', gap: 5, cursor: 'pointer' } },
            h('input', { type: 'checkbox', checked: unreadOnly, onChange: (ev) => { setUnreadOnly(ev.target.checked); setSel([]) } }),
            'Unread only'),
          h('select', { className: 'md-select', value: sort, onChange: (ev) => setSort(ev.target.value), title: 'Sort order' },
            h('option', { value: 'received' }, 'Newest'),
            h('option', { value: 'sender' }, 'Sender'),
            h('option', { value: 'subject' }, 'Subject')),
          h('select', { className: 'md-select', value: String(limit), onChange: (ev) => setLimit(Number(ev.target.value)), title: 'Page size' },
            h('option', { value: '50' }, '50'),
            h('option', { value: '100' }, '100'),
            h('option', { value: '200' }, '200')),
          h('button', {
            className: 'md-btn' + (askOpen ? ' primary' : ''),
            type: 'button',
            title: 'Ask the agent about this mailbox without leaving the panel',
            onClick: () => setAskOpen((v) => !v),
          }, askOpen ? 'Assistant on' : 'Ask'),
          h('button', { className: 'md-btn', type: 'button', disabled: busy !== '', onClick: () => void openCompose('new') }, 'New'),
          h('button', { className: 'md-btn', type: 'button', disabled: busy !== '', onClick: () => void refresh() }, busy === 'refresh' ? 'Refreshing…' : 'Refresh')),
        error !== '' && listState === 'error'
          ? h('div', { className: 'md-err', role: 'alert' },
              noOutlook
                ? h('div', null,
                    'No classic Outlook answered on this machine.',
                    h('div', { style: { marginTop: 8 } },
                      'This panel drives desktop Outlook over COM. If this PC uses the new Outlook (the WebView2 client), or Office is not installed, there is no COM surface to drive — install classic Outlook, or look at the interface with the sample mailbox.'),
                    h('div', { style: { marginTop: 10, display: 'flex', gap: 8, alignItems: 'center' } },
                      h('button', { className: 'md-btn primary', type: 'button', disabled: busy !== '', onClick: () => void switchSampleMail(true) }, busy === 'sample' ? 'Loading…' : 'Preview with sample mail'),
                      h('span', { style: { opacity: .7 } }, 'Synthetic mail, labelled SAMPLE MAIL — nothing is sent or saved.')),
                    h('div', { style: { marginTop: 8, opacity: .75, fontSize: 11 } }, error))
                : error)
          : null,
        h('div', { className: 'md-body' },
          rail,
          h('div', { className: 'md-list' },
            toolbar,
            h('div', { className: 'md-rows', ref: rowsRef, role: 'listbox', 'aria-label': 'Messages', 'aria-multiselectable': true }, rows),
            h('div', { className: 'md-pager' },
              searchRows !== null
                ? [
                    h('span', { key: 'found' }, searchBusy
                      ? 'Searching…'
                      : searchRows.length + ' result' + (searchRows.length === 1 ? '' : 's') + ' for "' + submittedSearch + '"'),
                    h('span', { key: 'sp', className: 'md-spacer' }),
                    h('button', { key: 'clear', className: 'md-btn', type: 'button', onClick: clearSearch }, 'Clear search'),
                  ]
                : null,
              // While the page is in flight the counters are still zeroed, and
              // "1–0 of 0 · unread count unavailable" reads like an empty
              // mailbox. Say what is actually happening.
              searchRows !== null ? null : listState === 'loading'
                ? [h('span', { key: 'loading' }, 'Loading…'), h('span', { key: 'sp', className: 'md-spacer' })]
                : [
                    h('span', { key: 'range' }, page.offset + 1 + '–' + (page.offset + page.returned) + ' of ' + page.total),
                    page.unreadInFolder > 0 ? h('span', { key: 'unread' }, '· ' + page.unreadInFolder + ' unread') : null,
                    page.unreadInFolder === -1 ? h('span', { key: 'unc' }, '· unread count unavailable') : null,
                    h('span', { key: 'sp', className: 'md-spacer' }),
                    page.offset + page.returned < page.total
                      ? h('button', { key: 'more', className: 'md-btn', type: 'button', disabled: busy !== '', onClick: () => void loadMore() }, busy === 'more' ? 'Loading…' : 'Load more')
                      : h('span', { key: 'end', style: { opacity: .6 } }, 'end of folder'),
                  ])),
          h('div', { className: 'md-read' }, reader)),
        askDrawer,
        moveDialog,
        composeView,
        toast !== null ? h('div', { className: 'md-toast', style: toast.isError ? { color: 'var(--dsw-alias-label-danger,#e5484d)' } : undefined, role: 'status' }, toast.text) : null)
    }

    /* ------------------------------------------------------------- wiring */

    /**
     * Chat-link click interception (kept from upstream v2.11.0): a plain click
     * on an anchor whose path is /olk-notify/open?entryId=… is answered with
     * the same same-origin POST the badge popover uses, so the mail opens in
     * Outlook with no extra browser tab.
     */
    const openMailByEntryId = async (entryId) => {
      try {
        const res = await fetch('/olk-notify/open', {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ entryId }),
        })
        const data = await res.json().catch(() => null)
        if (data && data.ok) toastHost('📨 Opened in Outlook: ' + (data.subject || '(no subject)'))
        else toastHost('⚠️ Could not open: ' + ((data && data.error) || ('HTTP ' + res.status)), true)
      } catch (error) {
        toastHost('⚠️ Could not open: ' + (error && error.message ? error.message : 'network error'), true)
      }
    }

    const interceptChatLink = (ev) => {
      if (ev.defaultPrevented || ev.button !== 0 || ev.metaKey || ev.ctrlKey || ev.shiftKey || ev.altKey) return
      const target = ev.target
      if (target === null || typeof target.closest !== 'function') return
      const a = target.closest('a[href]')
      if (a === null) return
      let u
      try { u = new URL(a.href, location.href) } catch { return }
      if (u.pathname !== '/olk-notify/open') return
      const entryId = u.searchParams.get('entryId')
      if (entryId === null || entryId === '') return
      ev.preventDefault()
      ev.stopPropagation()
      void openMailByEntryId(entryId)
    }

    /* ---- the draft card in chat --------------------------------------- */
    // The panel and the chat card are two views of one host-owned draft (see
    // the compose routes in index.js), so the only thing they must share is
    // "open this draft". A module-level bus keeps that out of the React trees
    // the harness owns.
    const draftBus = {
      listeners: new Set(),
      subscribe(fn) { this.listeners.add(fn); return () => { this.listeners.delete(fn) } },
      emit(draftId) { for (const fn of this.listeners) fn(draftId) },
    }
    /** Set by apply once ctx.layout is live: navigates the frame to the mail
     *  panel and asks it to open one draft. */
    let focusMailDraft = null
    /** The panel is not mounted while the chat is showing, so an "open this
     *  draft" request that arrives first has to survive until it mounts. */
    let pendingMailDraft = null
    /** Same idea for the right-pane tab: the id it should show, kept until the
     *  tab body mounts and reads its own navigation parameters. */
    let pendingRightDraft = null

    function snippetOf(html, max) {
      const text = String(html || '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim()
      return text.length > max ? text.slice(0, max - 1) + '…' : text
    }

    /** The tool half answers with a JSON text block, so the card reads the
     *  draftId back out of the result content. */
    function draftIdOf(block) {
      if (block === null || block === undefined) return ''
      const content = Array.isArray(block.content) ? block.content : []
      const text = content.filter((c) => c !== null && c !== undefined && c.type === 'text')
        .map((c) => String(c.text || '')).join('')
      if (text === '') return ''
      try {
        const parsed = JSON.parse(text)
        if (parsed !== null && typeof parsed === 'object' && typeof parsed.draftId === 'string') return parsed.draftId
      } catch { /* not a JSON block, fall back to the regex */ }
      const match = /"draftId"\s*:\s*"([^"]+)"/.exec(text)
      return match === null ? '' : match[1]
    }

    function DraftCard(props) {
      const phase = props === null || props === undefined ? '' : props.phase
      const block = props === null || props === undefined ? null : props.block
      const [draft, setDraft] = React.useState(null)
      const [state, setState] = React.useState('ready')
      const [error, setError] = React.useState('')
      const [busy, setBusy] = React.useState('')
      const [confirm, setConfirm] = React.useState(null)
      const settled = React.useRef(false)
      const draftId = phase === 'result' ? draftIdOf(block) : ''
      settled.current = state !== 'ready'

      // The harness has no host→client push channel we can rely on, so the
      // card polls until it is sent or discarded. Two seconds is cheap: the
      // route only touches an in-memory map.
      React.useEffect(() => {
        if (draftId === '') return undefined
        let live = true
        const load = async () => {
          if (settled.current) return
          const res = await api('/maildesk/api/draft', { op: 'get', draftId })
          if (!live || settled.current) return
          if (res.ok === true) { setDraft(res.draft); return }
          setState('gone')
          setError(res.error || 'this draft is no longer open')
        }
        void load()
        const timer = setInterval(() => { void load() }, 2000)
        return () => { live = false; clearInterval(timer) }
      }, [draftId])

      const send = async () => {
        setBusy('send')
        setError('')
        const res = await api('/maildesk/api/send', { draftId })
        const result = res.result || {}
        setBusy('')
        if (res.ok === true) { setState('sent'); setConfirm(null); return }
        if (result.status === 'needs-confirmation') { setConfirm({ hash: result.hash, preview: result.preview || {} }); return }
        setError(result.error || res.error || 'send failed')
      }

      const confirmSend = async () => {
        if (confirm === null) return
        setBusy('confirm')
        const res = await api('/maildesk/api/send', { draftId, confirmHash: confirm.hash })
        const result = res.result || {}
        setBusy('')
        if (res.ok === true) { setState('sent'); setConfirm(null); return }
        setConfirm(null)
        setError(result.status === 'needs-confirmation'
          ? 'The draft changed while you were confirming — send it again.'
          : (result.error || res.error || 'send failed'))
      }

      const discard = async () => {
        setBusy('discard')
        await api('/maildesk/api/draft', { op: 'discard', draftId })
        setBusy('')
        setState('discarded')
      }

      const shell = (stateClass, title, body, actions) => h('div', { className: 'md-card ' + stateClass },
        h('div', { className: 'md-card-head' },
          h('span', { className: 'md-card-badge' }, 'Draft'),
          h('span', { className: 'md-card-title' }, title)),
        body === null ? null : h('div', { className: 'md-card-body' }, body),
        actions === null ? null : h('div', { className: 'md-card-actions' }, actions))

      if (phase !== 'result') return shell('md-card-pending', 'Composing mail…', 'preparing the draft', null)
      if (draftId === '') return shell('md-card-pending', 'Draft', 'the tool did not return a draft id', null)
      if (state === 'sent') return shell('md-card-done', 'Sent', 'The message left the draft store and went to Outlook for delivery.', null)
      if (state === 'discarded') return shell('md-card-done', 'Draft discarded', null, null)
      if (state === 'gone') return shell('md-card-done', 'Draft closed', error === '' ? 'This draft is no longer open.' : error, null)
      if (draft === null) return shell('md-card-pending', 'Draft', 'reading it…', null)

      const to = formatPeople(draft.to)
      const cc = formatPeople(draft.cc)
      const subject = draft.subject === '' ? '(no subject)' : draft.subject
      const modeLabel = draft.mode === 'forward' ? 'Forward' : draft.mode === 'replyAll' ? 'Reply all' : draft.mode === 'reply' ? 'Reply' : 'New mail'
      const actions = confirm !== null
        ? [
            h('button', { className: 'md-btn md-btn-primary', onClick: confirmSend, disabled: busy !== '' },
              busy === 'confirm' ? 'Sending…' : 'Send now'),
            h('button', { className: 'md-btn', onClick: () => setConfirm(null), disabled: busy !== '' }, 'Back'),
          ]
        : [
            h('button', { className: 'md-btn', onClick: () => { if (focusMailDraft !== null) focusMailDraft(draftId) } }, 'Edit in panel'),
            h('button', { className: 'md-btn', onClick: () => { if (openDraftBesideChat !== null) openDraftBesideChat(draftId) } }, 'Open beside chat'),
            h('button', { className: 'md-btn md-btn-primary', onClick: send, disabled: busy !== '' },
              busy === 'send' ? 'Sending…' : 'Send…'),
            h('button', { className: 'md-btn', onClick: discard, disabled: busy !== '' }, 'Discard'),
          ]

      return h('div', { className: 'md-card' },
        h('div', { className: 'md-card-head' },
          h('span', { className: 'md-card-badge' }, modeLabel),
          h('span', { className: 'md-card-title' }, subject)),
        h('div', { className: 'md-card-row' },
          h('span', { className: 'md-card-label' }, 'To'),
          h('span', { className: 'md-card-value' }, to === '' ? '(addressed by Outlook)' : to)),
        cc === '' ? null : h('div', { className: 'md-card-row' },
          h('span', { className: 'md-card-label' }, 'Cc'),
          h('span', { className: 'md-card-value' }, cc)),
        h('div', { className: 'md-card-body' }, snippetOf(draft.bodyHtml, 320) || 'The body is still empty.'),
        error === '' ? null : h('div', { className: 'md-card-error' }, error),
        confirm === null ? null : h('div', { className: 'md-card-confirm' },
          h('div', { className: 'md-card-confirm-head' }, 'Send this?'),
          h('div', { className: 'md-card-confirm-row' },
            h('b', {}, 'To '), String(confirm.preview.to || to || '(addressed by Outlook)')),
          h('div', { className: 'md-card-confirm-row' },
            h('b', {}, 'Subject '), String(confirm.preview.subject || subject)),
          h('div', { className: 'md-card-confirm-note' },
            'Only this exact content is sent — editing it afterwards requires a new confirmation.')),
        h('div', { className: 'md-card-actions' }, actions))
    }

    // --- the draft beside the chat (right-pane tab) -----------------------
    // A right-Sidebar tab is two registrations: a static type in
    // ctx.sidebarRightTabs, and the keyed body under the type's own `id`. The
    // body receives `useTabInfo`, which is how it finds the draft id the
    // opener passed as `params`.
    const MAILDESK_TAB_ID = 'dsh-arya-maildesk:draft'
    const MAILDESK_TAB_KIND = 'maildesk-draft'
    /** Set by apply once ctx.sidebarRight is live. */
    let openDraftBesideChat = null

    function RightPaneDraft(props) {
      const useTabInfo = props === null || props === undefined ? undefined : props.useTabInfo
      const [draft, setDraft] = React.useState(null)
      const [state, setState] = React.useState('ready')
      const [error, setError] = React.useState('')
      const [busy, setBusy] = React.useState('')
      const [confirm, setConfirm] = React.useState(null)

      let info = null
      if (typeof useTabInfo === 'function') {
        try { info = useTabInfo() } catch { info = null }
      }
      const navigation = info === null || info === undefined || info.tab === undefined ? null : info.tab.navigation
      const params = navigation === null || navigation === undefined ? null : navigation.params
      const opened = params !== null && params !== undefined && typeof params.draftId === 'string' ? params.draftId : ''
      const draftId = opened !== '' ? opened : (pendingRightDraft === null ? '' : pendingRightDraft)
      const tabActions = info === null || info === undefined || info.tab === undefined ? null : info.tab.actions

      React.useEffect(() => {
        if (draftId === '') return undefined
        let live = true
        const load = async () => {
          const res = await api('/maildesk/api/draft', { op: 'get', draftId })
          if (!live) return
          if (res.ok === true) { setDraft(res.draft); return }
          setState('gone')
          setError(res.error || 'this draft is no longer open')
        }
        void load()
        const timer = setInterval(() => { void load() }, 2000)
        return () => { live = false; clearInterval(timer) }
      }, [draftId])

      const send = async () => {
        setBusy('send')
        setError('')
        const res = await api('/maildesk/api/send', { draftId })
        const result = res.result || {}
        setBusy('')
        if (res.ok === true) { setState('sent'); setConfirm(null); return }
        if (result.status === 'needs-confirmation') { setConfirm({ hash: result.hash, preview: result.preview || {} }); return }
        setError(result.error || res.error || 'send failed')
      }

      const confirmSend = async () => {
        if (confirm === null) return
        setBusy('confirm')
        const res = await api('/maildesk/api/send', { draftId, confirmHash: confirm.hash })
        const result = res.result || {}
        setBusy('')
        if (res.ok === true) { setState('sent'); setConfirm(null); return }
        setConfirm(null)
        setError(result.status === 'needs-confirmation'
          ? 'The draft changed while you were confirming — send it again.'
          : (result.error || res.error || 'send failed'))
      }

      const discard = async () => {
        setBusy('discard')
        await api('/maildesk/api/draft', { op: 'discard', draftId })
        setBusy('')
        setState('discarded')
      }

      if (draftId === '') {
        return h('div', { className: 'md-rt' },
          h('p', { className: 'md-rt-note' }, 'No draft yet. Ask the agent for a reply, or open one from the mail panel.'))
      }
      if (state === 'sent') return h('div', { className: 'md-rt' }, h('p', { className: 'md-rt-note' }, 'Sent for delivery.'))
      if (state === 'discarded') return h('div', { className: 'md-rt' }, h('p', { className: 'md-rt-note' }, 'Draft discarded.'))
      if (state === 'gone') return h('div', { className: 'md-rt' }, h('p', { className: 'md-rt-note' }, error === '' ? 'This draft is no longer open.' : error))
      if (draft === null) return h('div', { className: 'md-rt' }, h('p', { className: 'md-rt-note' }, 'Reading the draft…'))

      const to = formatPeople(draft.to)
      const cc = formatPeople(draft.cc)
      const subject = draft.subject === '' ? '(no subject)' : draft.subject
      const modeLabel = draft.mode === 'forward' ? 'Forward' : draft.mode === 'replyAll' ? 'Reply all' : draft.mode === 'reply' ? 'Reply' : 'New mail'
      const actions = confirm !== null
        ? [
            h('button', { className: 'md-btn md-btn-primary', onClick: confirmSend, disabled: busy !== '' },
              busy === 'confirm' ? 'Sending…' : 'Send now'),
            h('button', { className: 'md-btn', onClick: () => setConfirm(null), disabled: busy !== '' }, 'Back'),
          ]
        : [
            h('button', { className: 'md-btn md-btn-primary', onClick: send, disabled: busy !== '' },
              busy === 'send' ? 'Sending…' : 'Send…'),
            h('button', { className: 'md-btn', onClick: () => { if (focusMailDraft !== null) focusMailDraft(draftId) } }, 'Open in panel'),
            h('button', { className: 'md-btn', onClick: discard, disabled: busy !== '' }, 'Discard'),
          ]

      return h('div', { className: 'md-rt' },
        h('div', { className: 'md-rt-head' },
          h('span', { className: 'md-card-badge' }, modeLabel),
          h('span', { className: 'md-rt-title' }, subject)),
        h('div', { className: 'md-card-row' },
          h('span', { className: 'md-card-label' }, 'To'),
          h('span', { className: 'md-card-value' }, to === '' ? '(addressed by Outlook)' : to)),
        cc === '' ? null : h('div', { className: 'md-card-row' },
          h('span', { className: 'md-card-label' }, 'Cc'),
          h('span', { className: 'md-card-value' }, cc)),
        h('div', { className: 'md-rt-body' }, snippetOf(draft.bodyHtml, 1200) || 'The body is still empty.'),
        error === '' ? null : h('div', { className: 'md-card-error' }, error),
        confirm === null ? null : h('div', { className: 'md-card-confirm' },
          h('div', { className: 'md-card-confirm-head' }, 'Send this?'),
          h('div', { className: 'md-card-confirm-row' },
            h('b', {}, 'To '), String(confirm.preview.to || to || '(addressed by Outlook)')),
          h('div', { className: 'md-card-confirm-row' },
            h('b', {}, 'Subject '), String(confirm.preview.subject || subject)),
          h('div', { className: 'md-card-confirm-note' },
            'Only this exact content is sent — editing it afterwards requires a new confirmation.')),
        h('div', { className: 'md-card-actions' }, actions),
        tabActions === null ? null : h('div', { className: 'md-rt-foot' },
          h('button', { className: 'md-link', onClick: () => { try { tabActions.close() } catch { /* tab already gone */ } } }, 'Close this tab')))
    }

    return {
      name: 'dsh-arya-maildesk',
      inject: ['slots', 'layout', 'sidebarRightTabs', 'sidebarRight'],
      apply(ctx) {
        beacon('apply called (mail panel + badge)')
        const stylesTag = ensureStyles()
        if (stylesTag !== null) {
          ctx.effect(() => () => {
            if (stylesTag.parentNode !== null) stylesTag.parentNode.removeChild(stylesTag)
          }, 'dsh-arya-maildesk: styles')
        }
        document.addEventListener('click', interceptChatLink, true)
        ctx.effect(() => () => {
          document.removeEventListener('click', interceptChatLink, true)
          for (const div of document.querySelectorAll('.md-toast')) {
            if (div.parentNode !== null) div.parentNode.removeChild(div)
          }
        }, 'dsh-arya-maildesk: chat link interception')

        // The mailbox page, selected by its sidebar icon.
        ctx.slots.inject('main', () => ctx.slots.register({ name: 'main', key: 'mail' }, Panel))
        // The draft the agent opens: one keyed row view for maildesk_compose.
        // A keyed tool.call.toolview entry replaces the generic row for that
        // exact tool name and leaves every other tool row alone.
        ctx.slots.inject('tool.call.toolview', () => ctx.slots.register(
          { name: 'tool.call.toolview', key: 'maildesk_compose' },
          DraftCard,
        ))
        // [Edit in panel]: select the mail panel, then hand it the draft.
        // selectPanel throws when the key is not registered, so a host without
        // our main panel still keeps the card usable.
        focusMailDraft = (draftId) => {
          pendingMailDraft = draftId
          try { ctx.layout.selectPanel('mail') } catch { /* no mail panel here */ }
          draftBus.emit(draftId)
        }
        // [Open beside chat]: the draft in the right Sidebar, next to the
        // conversation instead of replacing it. A tab is a static type plus a
        // keyed body, and only `openTab` after both exist.
        ctx.effect(() => ctx.sidebarRightTabs.register({
          id: MAILDESK_TAB_ID,
          kind: MAILDESK_TAB_KIND,
          title: () => 'Mail draft',
        }), 'dsh-arya-maildesk: draft tab type')
        ctx.slots.inject('sidebar.right.pane.tab', () => ctx.slots.register(
          { name: 'sidebar.right.pane.tab', key: MAILDESK_TAB_ID },
          RightPaneDraft,
        ))
        openDraftBesideChat = (draftId) => {
          pendingRightDraft = draftId
          try {
            ctx.sidebarRight.openTab(MAILDESK_TAB_KIND, { params: { draftId } })
          } catch (err) {
            console.warn('[dsh-arya-maildesk] the draft tab could not open', err)
          }
        }
        ctx.slots.inject('sidebar.panellist', () => ctx.slots.register(
          { name: 'sidebar.panellist', id: 'mail', order: -6, label: 'Mail' },
          (props) => h(MailIcon, { size: props && props.wide === false ? 18 : 16 }),
        ))
        // The doorbell: new arrivals only.
        ctx.slots.inject('sidebar.footer.action', () => ctx.slots.register(
          { name: 'sidebar.footer.action', id: 'newmail-badge', order: 10, label: 'New mail' },
          (props) => h(NewMailRow, props),
        ))
      },
    }
  },
})

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
  id: 'dsh-arya-maildesk',
  factory: (require) => {
    const React = require('react')
    const ReactDOM = require('react-dom')
    const h = React.createElement

    const STATE_URL = '/olk-notify/api/state'
    const CALL_URL = '/maildesk/api/call'
    const HEALTH_URL = '/maildesk/api/health'
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

        /* --- main panel --- */
        '.md{display:flex;flex-direction:column;height:100%;min-height:0;font:13px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif;color:var(--dsw-alias-label-primary,#1b1b1b);background:var(--dsh-bg,#fff)}',
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
        '.md-empty{flex:1 1 auto;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;padding:28px;text-align:center;font-size:12.5px;opacity:.8}',
        '.md-empty-h{font-size:13.5px;font-weight:600;opacity:1}',
        '.md-err{margin:10px 14px;padding:10px 12px;border-radius:10px;background:#e5484d12;color:#a3262b;font-size:12px;line-height:1.55;white-space:pre-wrap;word-break:break-word}',
        '.md-skel{padding:7px 10px;border-bottom:1px solid #f2f2f2}',
        '.md-skel div{height:9px;border-radius:5px;background:#8898ac1f;margin:4px 0}',
        '.md-skel div:nth-child(1){width:34%}.md-skel div:nth-child(2){width:78%}.md-skel div:nth-child(3){width:60%;height:8px}',
        '.md-pager{display:flex;align-items:center;gap:8px;padding:8px 10px;border-top:1px solid var(--dsw-alias-border-subtle,#f0f0f0);font-size:11.5px;opacity:.8;flex:none}',
        '.md-kbd{font-size:10.5px;padding:1px 5px;border-radius:4px;border:1px solid var(--dsw-alias-border-subtle,#d9d9d9);background:#8898ac10}',
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
      const [folderId, setFolderId] = React.useState(null)
      const [page, setPage] = React.useState({ total: 0, offset: 0, returned: 0, folderPath: '', mails: [], unreadInFolder: -1 })
      const [listState, setListState] = React.useState('loading')
      const [sel, setSel] = React.useState([])
      const [openId, setOpenId] = React.useState(null)
      const [openMail, setOpenMail] = React.useState(null)
      const [readState, setReadState] = React.useState('idle')
      const [error, setError] = React.useState('')
      const [busy, setBusy] = React.useState('')
      const [toast, setToast] = React.useState(null)
      const [unreadOnly, setUnreadOnly] = React.useState(false)
      const [sort, setSort] = React.useState('received')
      const [limit, setLimit] = React.useState(100)
      const [provider, setProvider] = React.useState('com')
      const [moveTo, setMoveTo] = React.useState('')
      const anchor = React.useRef(null)
      const rowsRef = React.useRef(null)

      const announce = (text, isError) => {
        setToast({ text, isError: isError === true })
        setTimeout(() => setToast(null), 2800)
      }

      React.useEffect(() => {
        void (async () => {
          try {
            const res = await fetch(HEALTH_URL)
            const data = await res.json()
            if (data && data.ok === true) setProvider(data.provider === 'fixture' ? 'fixture' : 'com')
          } catch { /* keep the com default */ }
          try {
            const result = await call('list_folders', { maxDepth: 6 })
            const list = Array.isArray(result.folders) ? result.folders : []
            setFolders(list)
            const first = list.find((f) => /^inbox$/i.test(f.name)) || list[0]
            if (first) setFolderId(first.entryId)
            else setListState('empty')
          } catch (error) {
            setError(error.message)
            setFolders([])
            setListState('error')
          }
        })()
      }, [])

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
          const result = await call('read_mail', { entryId, maxChars: 20000 })
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

      const doMove = async () => {
        const ids = targetIds()
        if (ids.length === 0 || moveTo === '') return
        setBusy('move')
        try {
          const result = await call('move_mail', { entryIds: ids, destFolder: moveTo })
          if (Array.isArray(result.failed) && result.failed.length > 0) {
            announce(result.failed.length + ' could not be moved: ' + result.failed[0].error, true)
          } else {
            announce('Moved ' + (result.moved || ids.length) + ' to ' + (result.destPath || 'folder'))
          }
          setSel([])
          setMoveTo('')
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

      const onKeyDown = (ev) => {
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

      const rail = h('div', { className: 'md-rail', role: 'navigation', 'aria-label': 'Folders' },
        folders === null ? [0, 1, 2, 3].map((i) => h('div', { key: i, className: 'md-skel' }, h('div'), h('div'), h('div'))) : null,
        folders !== null ? (() => {
          const out = []
          let store = null
          for (const f of folders) {
            if (f.store !== store) {
              store = f.store
              out.push(h('div', { key: 'store-' + store, className: 'md-rail-store' }, store))
            }
            const unread = typeof f.unread === 'number' ? f.unread : 0
            out.push(h('button', {
              key: f.entryId,
              type: 'button',
              className: 'md-folder' + (f.entryId === folderId ? ' on' : ''),
              style: { paddingLeft: 8 + (f.depth || 0) * 12 + 'px' },
              title: f.path,
              onClick: () => { setFolderId(f.entryId); setSel([]); setOpenId(null); setOpenMail(null); setReadState('idle') },
            },
              h(FolderIcon, { size: 13 }),
              h('span', { className: 'md-folder-name' }, f.name),
              unread > 0 ? h('span', { className: 'md-unread' }, String(unread)) : null,
              h('span', { className: 'md-count' }, String(typeof f.itemCount === 'number' ? f.itemCount : ''))))
          }
          return out
        })() : null,
        h('div', { style: { padding: '10px 8px 2px', fontSize: '10.5px', opacity: .6 } },
          'Badge counts new arrivals only.'),
      )

      const toolbar = h('div', { className: 'md-toolbar' },
        h('button', { className: 'md-btn primary', type: 'button', disabled: busy !== '' || selectionCount === 0, onClick: () => void doMark(true) }, 'Mark read'),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || selectionCount === 0, onClick: () => void doMark(false) }, 'Mark unread'),
        h('button', { className: 'md-btn danger', type: 'button', disabled: busy !== '' || selectionCount === 0, onClick: () => void doDelete() }, 'Delete'),
        h('select', {
          className: 'md-select',
          value: moveTo,
          disabled: busy !== '' || selectionCount === 0,
          title: 'Move the selection to a folder',
          onChange: (ev) => setMoveTo(ev.target.value),
        },
          h('option', { value: '' }, 'Move to…'),
          (folders || []).map((f) => h('option', { key: f.entryId, value: f.entryId }, '  '.repeat(f.depth || 0) + f.name))),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || moveTo === '' || selectionCount === 0, onClick: () => void doMove() }, 'Go'),
        h('button', { className: 'md-btn', type: 'button', disabled: busy !== '' || selectionCount === 0, onClick: () => void doOpenInOutlook() }, 'Open in Outlook'),
        h('span', { className: 'md-selinfo' }, selectionCount > 0 ? selectionCount + ' selected' : 'no selection'))

      const rows = (() => {
        if (listState === 'loading') {
          return [0, 1, 2, 3, 4, 5, 6].map((i) => h('div', { key: i, className: 'md-skel' }, h('div'), h('div'), h('div')))
        }
        if (listState === 'error') return null
        if (page.mails.length === 0) {
          return h('div', { className: 'md-empty' },
            h('div', { className: 'md-empty-h' }, unreadOnly ? 'No unread mail here' : 'This folder is empty'),
            h('div', null, page.folderPath || ''))
        }
        return page.mails.map((m) => h('div', {
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
        return h(React.Fragment, null,
          h('div', { className: 'md-read-head' },
            h('div', { className: 'md-read-subj' }, openMail.subject || '(no subject)'),
            h('div', { className: 'md-meta' },
              h('div', null, h('b', null, 'From: '), openMail.sender || ''),
              h('div', null, h('b', null, 'To: '), openMail.to || ''),
              openMail.cc ? h('div', null, h('b', null, 'Cc: '), openMail.cc) : null,
              h('div', null, h('b', null, 'Received: '), openMail.received || '')),
            attachments.length > 0 ? h('div', null, attachments.map((a, i) => h('span', { key: i, className: 'md-att' }, '📎 ' + a))) : null),
          h('div', { className: 'md-read-body' }, openMail.body || '(no text body)'))
      })()

      return h('div', { className: 'md', onKeyDown, tabIndex: -1 },
        h('div', { className: 'md-top' },
          h('span', { className: 'md-title' }, 'Mail'),
          provider === 'fixture'
            ? h('span', { className: 'md-pill fixture', title: 'MAILDESK_FIXTURE=1 — this is synthetic mail, not a mailbox' }, 'FIXTURE DATA')
            : h('span', { className: 'md-pill' }, 'Classic Outlook · COM'),
          page.folderPath ? h('span', { className: 'md-pill' }, page.folderPath) : null,
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
          h('button', { className: 'md-btn', type: 'button', disabled: busy !== '', onClick: () => void refresh() }, busy === 'refresh' ? 'Refreshing…' : 'Refresh')),
        error !== '' && listState === 'error'
          ? h('div', { className: 'md-err', role: 'alert' },
              noOutlook
                ? 'No classic Outlook answered on this machine.\n\nThis panel drives desktop Outlook over COM. If this PC uses the new Outlook (the WebView2 client), or Office is not installed, there is no COM surface to drive — install classic Outlook, or set MAILDESK_FIXTURE=1 to work on the interface with synthetic mail.\n\n' + error
                : error)
          : null,
        h('div', { className: 'md-body' },
          rail,
          h('div', { className: 'md-list' },
            toolbar,
            h('div', { className: 'md-rows', ref: rowsRef, role: 'listbox', 'aria-label': 'Messages', 'aria-multiselectable': true }, rows),
            h('div', { className: 'md-pager' },
              h('span', null, page.offset + 1 + '–' + (page.offset + page.returned) + ' of ' + page.total),
              page.unreadInFolder > 0 ? h('span', null, '· ' + page.unreadInFolder + ' unread') : null,
              page.unreadInFolder === -1 ? h('span', null, '· unread count unavailable') : null,
              h('span', { className: 'md-spacer' }),
              page.offset + page.returned < page.total
                ? h('button', { className: 'md-btn', type: 'button', disabled: busy !== '', onClick: () => void loadMore() }, busy === 'more' ? 'Loading…' : 'Load more')
                : h('span', { style: { opacity: .6 } }, 'end of folder'))),
          h('div', { className: 'md-read' }, reader)),
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

    return {
      name: 'dsh-arya-maildesk',
      inject: ['slots'],
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

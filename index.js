<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
<meta name="theme-color" content="#0d5c2f">
<title>PG Docs Sign</title>
<link rel="manifest" href="/manifest.json">
<link rel="icon" href="/SignDocs_icon.png">
<link rel="apple-touch-icon" href="/SignDocs_icon.png">
<style>
  :root {
    --green-900:#064e3b; --green-700:#0d5c2f; --green-500:#10b981; --green-100:#d1fadf; --green-50:#f0fdf4;
    --gold:#c9a227; --gold-light:#f7e08a;
    --bg:#f6f9f6; --card:#fff; --text:#0f172a; --muted:#64748b; --border:#e5e7eb;
    --red:#dc2626; --amber:#f59e0b; --blue:#2563eb;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
  html, body { overscroll-behavior-y: contain; overscroll-behavior-x: contain; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; -webkit-font-smoothing: antialiased; }
  .container { max-width: 1200px; margin: 0 auto; padding: 16px; }
  .center { display: flex; align-items: center; justify-content: center; min-height: 100vh; padding: 12px; }
  /* Corporate header - dark green, clean */
  .header { background: linear-gradient(180deg, var(--green-900) 0%, var(--green-700) 100%); border-bottom: 3px solid var(--gold); padding: 12px 18px; box-shadow: 0 4px 12px rgba(6,78,59,0.25); position: relative; }
  .header-row1 { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
  .brand { display: flex; align-items: center; gap: 12px; }
  .brand img { width: 48px; height: 48px; border-radius: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.35); border: 2px solid var(--gold); background: #fff; object-fit: cover; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.15; }
  .brand-text .title { font-size: 20px; font-weight: 800; color: #fff; letter-spacing: .5px; text-shadow: 0 1px 2px rgba(0,0,0,0.2); }
  .brand-text .subtitle { font-size: 11px; color: var(--gold-light); font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; margin-top: 2px; }
  .header-user { margin-top: 6px; font-size: 13px; color: #fff; display: flex; flex-direction: column; gap: 2px; align-items: flex-start; }
  .header-user .uname { font-weight: 700; font-size: 15px; color: #fff; display: flex; align-items: center; gap: 8px; cursor: pointer; }
  .header-user .uname svg { flex-shrink: 0; }
  .header-user .uname:hover .edit-icon { opacity: 1; }
  .header-user .edit-icon { opacity: 0.5; font-size: 11px; color: var(--gold-light); }
  .header-user .ustore { font-size: 12px; color: var(--gold-light); font-weight: 600; margin-left: 26px; }
  .menu-btn { background: rgba(255,255,255,0.15); color: #fff; border: 1px solid rgba(255,255,255,0.3); width: 40px; height: 40px; border-radius: 20px; padding: 0; font-size: 18px; cursor: pointer; display: flex; align-items: center; justify-content: center; }
  .menu-btn:hover { background: rgba(255,255,255,0.25); }
  .menu-drop { position: absolute; top: 100%; right: 12px; margin-top: 6px; background: #fff; border-radius: 10px; box-shadow: 0 6px 20px rgba(0,0,0,0.15); overflow: hidden; z-index: 100; min-width: 170px; border: 1px solid var(--border); }
  .menu-drop button { display: flex; width: 100%; padding: 12px 16px; background: #fff; color: var(--text); border: none; text-align: left; font-size: 14px; align-items: center; gap: 10px; cursor: pointer; }
  .menu-drop button:hover { background: var(--green-50); color: var(--green-900); }
  .menu-drop hr { border: none; border-top: 1px solid var(--border); margin: 0; }
  /* Card */
  .card { background: var(--card); border-radius: 12px; padding: 20px; box-shadow: 0 2px 12px rgba(6,78,59,0.06); border: 1px solid rgba(13,92,47,0.08); }
  .card h1 { font-size: 20px; margin-bottom: 6px; color: var(--green-900); }
  .card .subtitle { color: var(--muted); font-size: 13px; margin-bottom: 20px; }
  .login-card { width: 100%; max-width: 440px; }
  .login-header { text-align: center; margin-bottom: 20px; }
  .login-header img { width: 80px; height: 80px; border-radius: 40px; box-shadow: 0 4px 16px rgba(13,92,47,0.2); margin-bottom: 12px; }
  .login-header h1 { color: var(--green-900); font-size: 24px; font-weight: 800; letter-spacing: .5px; }
  .login-header .tag { color: var(--gold); font-size: 12px; font-weight: 700; letter-spacing: 2px; text-transform: uppercase; }
  /* Form */
  .form-group { margin-bottom: 14px; }
  label { display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px; color: #374151; }
  input, select, textarea { width: 100%; padding: 11px 12px; border: 1px solid #d1d5db; border-radius: 8px; font-size: 14px; font-family: inherit; background: #fff; transition: all .15s; }
  input:focus, select:focus, textarea:focus { outline: none; border-color: var(--green-700); box-shadow: 0 0 0 3px rgba(13,92,47,0.1); }
  button, .btn { background: var(--green-700); color: #fff; border: none; padding: 10px 18px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all .15s; }
  button:hover { background: var(--green-900); transform: translateY(-1px); box-shadow: 0 3px 8px rgba(13,92,47,0.2); }
  button.secondary { background: #f1f5f9; color: #334155; }
  button.secondary:hover { background: #e2e8f0; box-shadow: none; }
  button.danger { background: var(--red); }
  button.danger:hover { background: #b91c1c; }
  button.gold { background: linear-gradient(135deg, var(--gold) 0%, #a5811d 100%); }
  button:disabled { background: #9ca3af; cursor: not-allowed; transform: none; box-shadow: none; }
  .btn-full { width: 100%; padding: 12px; font-size: 15px; }
  .link { color: var(--green-700); text-decoration: none; font-size: 13px; cursor: pointer; font-weight: 600; }
  .link:hover { text-decoration: underline; }
  .msg { padding: 11px 14px; border-radius: 8px; font-size: 13px; margin-bottom: 12px; }
  .msg.error { background: #fef2f2; color: #b91c1c; border: 1px solid #fecaca; }
  .msg.success { background: #f0fdf4; color: #065f46; border: 1px solid #bbf7d0; }
  .msg.info { background: #eff6ff; color: #1e40af; border: 1px solid #bfdbfe; }
  .msg.warn { background: #fffbeb; color: #92400e; border: 1px solid #fde68a; }
  /* KPI cards */
  .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 20px; }
  .kpi { background: var(--card); border-radius: 12px; padding: 16px; border-left: 4px solid var(--green-700); box-shadow: 0 2px 8px rgba(6,78,59,0.05); }
  .kpi.gold { border-left-color: var(--gold); }
  .kpi.red { border-left-color: var(--red); }
  .kpi.blue { border-left-color: var(--blue); }
  .kpi .label { font-size: 11px; color: var(--muted); text-transform: uppercase; font-weight: 600; letter-spacing: .5px; }
  .kpi .value { font-size: 26px; font-weight: 800; color: var(--green-900); margin-top: 4px; }
  .kpi .sub { font-size: 12px; color: var(--muted); margin-top: 2px; }
  /* Tabs */
  .tabs { display: flex; gap: 2px; border-bottom: 2px solid var(--border); margin-bottom: 16px; overflow-x: auto; background: #fff; border-radius: 10px 10px 0 0; padding: 0 8px; }
  .tab { padding: 12px 18px; cursor: pointer; border-bottom: 3px solid transparent; font-size: 14px; color: var(--muted); white-space: nowrap; font-weight: 500; margin-bottom: -2px; }
  .tab.active { color: var(--green-700); border-color: var(--green-700); font-weight: 700; }
  .tab:hover:not(.active) { color: var(--green-900); background: var(--green-50); }
  /* Table with horizontal scroll wrapper */
  .table-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; max-width: 100%; border-radius: 10px; box-shadow: 0 2px 8px rgba(6,78,59,0.05); margin-bottom: 12px; background: var(--card); }
  .table-scroll::-webkit-scrollbar { height: 8px; }
  .table-scroll::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
  .table-scroll::-webkit-scrollbar-thumb { background: var(--green-700); border-radius: 4px; }
  .table-scroll::-webkit-scrollbar-thumb:hover { background: var(--green-900); }
  table { width: 100%; border-collapse: collapse; background: var(--card); min-width: 600px; }
  th, td { padding: 12px 14px; text-align: left; font-size: 14px; vertical-align: top; }
  th { background: var(--green-900); color: #fff; font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: .5px; }
  tr { border-top: 1px solid #f1f5f9; }
  tr:hover td { background: var(--green-50); }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; text-transform: uppercase; }
  .badge.pending { background: #fef3c7; color: #92400e; }
  .badge.approved { background: #d1fae5; color: #065f46; }
  .badge.rejected { background: #fee2e2; color: #991b1b; }
  .badge.urgent { background: #fee2e2; color: #991b1b; }
  .badge.high { background: #ffedd5; color: #9a3412; }
  .badge.normal { background: #dbeafe; color: #1e40af; }
  .badge.low { background: #f3f4f6; color: #374151; }
  .age-badge { display: inline-block; padding: 3px 8px; border-radius: 10px; font-size: 11px; font-weight: 700; }
  .age-badge.fresh { background: #d1fadf; color: #065f46; }
  .age-badge.warm { background: #fef3c7; color: #92400e; }
  .age-badge.hot { background: #fed7aa; color: #9a3412; }
  .age-badge.critical { background: #fecaca; color: #991b1b; }
  .actions { display: flex; gap: 6px; flex-wrap: wrap; }
  .actions button { padding: 6px 12px; font-size: 12px; }
  .empty { text-align: center; padding: 40px; color: #9ca3af; font-size: 14px; }
  /* Multi-store picker */
  .store-picker { max-height: 210px; overflow-y: auto; border: 1px solid #d1d5db; border-radius: 8px; padding: 6px; background: #fff; }
  .store-picker::-webkit-scrollbar { width: 8px; }
  .store-picker::-webkit-scrollbar-thumb { background: var(--green-700); border-radius: 4px; }
  .store-opt { display: flex; align-items: center; gap: 10px; padding: 8px 10px; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .store-opt:hover { background: var(--green-50); }
  .store-opt input { width: auto; margin: 0; flex-shrink: 0; width: 18px; height: 18px; accent-color: var(--green-700); }
  .store-opt span { color: var(--text); }
  .store-chips { display: flex; flex-wrap: wrap; gap: 4px; }
  .store-chip { background: var(--green-100); color: var(--green-900); border: 1px solid var(--green-500); border-radius: 10px; padding: 2px 8px; font-size: 11px; font-weight: 600; white-space: nowrap; }
  /* Remarks */
  .remarks-cell { min-width: 220px; max-width: 300px; }
  .remark-box { background: #fffbeb; border: 1px solid #fde68a; border-left: 3px solid var(--gold); border-radius: 6px; padding: 8px 10px; }
  .remark-text { font-size: 13px; color: #78350f; white-space: pre-wrap; word-break: break-word; }
  .remark-meta { font-size: 10px; color: #a16207; margin-top: 4px; }
  .remark-actions { display: flex; gap: 4px; margin-top: 6px; }
  .remark-actions button { padding: 3px 8px; font-size: 11px; }
  .staff-remark { background: #fffbeb; border: 1px solid #fde68a; border-left: 3px solid var(--gold); border-radius: 6px; padding: 8px 10px; margin-top: 6px; font-size: 12px; color: #78350f; white-space: pre-wrap; }
  /* Guide */
  .guide-hero { background: linear-gradient(135deg, var(--green-900), var(--green-700)); border-radius: 14px; padding: 22px; display: flex; gap: 18px; align-items: center; margin-bottom: 20px; border-bottom: 4px solid var(--gold); box-shadow: 0 4px 16px rgba(6,78,59,0.2); }
  .guide-hero img { width: 64px; height: 64px; border-radius: 32px; border: 3px solid var(--gold); background: #fff; flex-shrink: 0; }
  .guide-hero h2 { color: #fff; font-size: 22px; font-weight: 800; letter-spacing: .3px; }
  .guide-hero p { color: var(--gold-light); font-size: 13px; font-weight: 600; margin-top: 4px; }
  .guide-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 14px; margin-bottom: 20px; }
  .guide-card { background: var(--card); border-radius: 12px; padding: 16px 18px; box-shadow: 0 2px 10px rgba(6,78,59,0.06); border-left: 4px solid var(--green-700); }
  .guide-card-head { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; flex-wrap: wrap; }
  .guide-num { background: var(--green-700); color: #fff; width: 26px; height: 26px; border-radius: 13px; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 13px; flex-shrink: 0; }
  .guide-icon { font-size: 20px; }
  .guide-title { font-weight: 700; color: var(--green-900); font-size: 15px; }
  .guide-body { font-size: 13.5px; color: #374151; line-height: 1.6; }
  .guide-compare { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 20px; }
  .guide-good, .guide-bad { border-radius: 12px; padding: 16px 18px; }
  .guide-good { background: #f0fdf4; border: 1px solid #bbf7d0; }
  .guide-bad { background: #fef2f2; border: 1px solid #fecaca; }
  .guide-good h4 { color: #065f46; margin-bottom: 10px; font-size: 15px; }
  .guide-bad h4 { color: #991b1b; margin-bottom: 10px; font-size: 15px; }
  .guide-good ul, .guide-bad ul { margin-left: 18px; font-size: 13px; line-height: 1.9; }
  .guide-good li { color: #065f46; }
  .guide-bad li { color: #991b1b; }
  .guide-steps { background: var(--card); border-radius: 12px; padding: 18px; box-shadow: 0 2px 10px rgba(6,78,59,0.06); margin-bottom: 20px; border-left: 4px solid var(--gold); }
  .guide-steps h4 { color: var(--green-900); margin-bottom: 12px; font-size: 15px; }
  .guide-steps ol { margin-left: 20px; font-size: 13.5px; line-height: 2; color: #374151; }
  .guide-status { background: var(--card); border-radius: 12px; padding: 18px; box-shadow: 0 2px 10px rgba(6,78,59,0.06); margin-bottom: 20px; }
  .guide-status h4 { color: var(--green-900); margin-bottom: 12px; font-size: 15px; }
  .guide-status-row { display: flex; gap: 12px; align-items: flex-start; padding: 8px 0; border-top: 1px solid #f1f5f9; font-size: 13px; color: #374151; }
  .guide-status-row:first-of-type { border-top: none; }
  .guide-status-row > span:first-child { flex-shrink: 0; min-width: 76px; }
  .guide-footer { background: linear-gradient(135deg, var(--gold), #b58f1c); color: #1a1a1a; border-radius: 12px; padding: 16px 20px; font-weight: 700; text-align: center; font-size: 14px; margin-bottom: 20px; }
  @media (max-width: 720px) {
    .guide-compare { grid-template-columns: 1fr; }
    .guide-hero { padding: 16px; gap: 12px; }
    .guide-hero img { width: 48px; height: 48px; }
    .guide-hero h2 { font-size: 18px; }
  }
  .box { background: var(--card); border-radius: 12px; padding: 20px; box-shadow: 0 2px 12px rgba(6,78,59,0.06); border: 1px solid rgba(13,92,47,0.08); margin-bottom: 20px; }
  .box h3 { font-size: 16px; margin-bottom: 12px; color: var(--green-900); display: flex; align-items: center; gap: 8px; }
  .box h3::before { content: ''; width: 4px; height: 18px; background: var(--gold); border-radius: 2px; }
  .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  .file-drop { border: 2px dashed var(--green-500); border-radius: 10px; padding: 20px; text-align: center; color: var(--green-700); cursor: pointer; background: var(--green-50); font-weight: 600; }
  .file-drop:hover { background: var(--green-100); border-color: var(--green-700); }
  .file-drop.has-file { border-style: solid; }
  .file-list { margin-top: 8px; display: flex; flex-direction: column; gap: 4px; }
  .file-item { display: flex; justify-content: space-between; align-items: center; background: #f8fafc; padding: 6px 10px; border-radius: 6px; font-size: 12px; }
  .file-item .rm { color: var(--red); cursor: pointer; font-weight: 700; }
  #fileInput, #cameraInput { display: none; }
  textarea { min-height: 60px; resize: vertical; }
  .upload-note { background: #fffbeb; border: 1px solid #fde68a; border-radius: 8px; padding: 10px 14px; font-size: 12px; color: #92400e; margin-bottom: 12px; }
  .upload-note strong { color: #78350f; }
  .upload-buttons { display: flex; gap: 8px; margin-top: 6px; }
  .upload-buttons button { flex: 1; padding: 10px; font-size: 13px; }
  /* Folder groups */
  .folder-group { background: var(--card); border-radius: 10px; margin-bottom: 12px; border: 1px solid var(--border); overflow: hidden; }
  .folder-head { padding: 12px 16px; background: var(--green-50); cursor: pointer; display: flex; justify-content: space-between; align-items: center; }
  .folder-head:hover { background: var(--green-100); }
  .folder-head .name { font-weight: 700; color: var(--green-900); font-size: 14px; }
  .folder-head .count { background: var(--green-700); color: #fff; padding: 2px 10px; border-radius: 10px; font-size: 12px; font-weight: 600; }
  .folder-body { display: none; padding: 0; }
  .folder-group.open .folder-body { display: block; }
  .folder-group.open .folder-head { border-bottom: 1px solid var(--border); }
  /* Signing modal - bottom toolbar */
  .modal { position: fixed; inset: 0; background: rgba(6,78,59,.92); z-index: 1000; display: flex; flex-direction: column; }
  .modal-topbar { background: var(--green-900); color: #fff; padding: 10px 14px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 6px rgba(0,0,0,0.3); }
  .modal-topbar .title { font-size: 14px; font-weight: 700; }
  .modal-topbar button { background: rgba(255,255,255,0.15); color: #fff; border: 1px solid rgba(255,255,255,0.3); padding: 6px 12px; font-size: 12px; }
  .modal-doc { flex: 1; overflow: auto; padding: 20px 20px 100px; display: flex; flex-direction: column; align-items: center; gap: 12px; }
  .modal-bottombar { background: #fff; padding: 8px; display: flex; justify-content: flex-start; gap: 4px; box-shadow: 0 -4px 12px rgba(0,0,0,0.15); overflow-x: auto; -webkit-overflow-scrolling: touch; scroll-snap-type: x mandatory; }
  .modal-bottombar::-webkit-scrollbar { height: 4px; }
  .modal-bottombar::-webkit-scrollbar-thumb { background: var(--green-700); border-radius: 2px; }
  .modal-tool-btn { display: flex; flex-direction: column; align-items: center; gap: 4px; padding: 8px 12px; background: transparent; color: var(--green-900); border: 1px solid transparent; border-radius: 8px; font-size: 11px; font-weight: 600; min-width: 64px; cursor: pointer; flex-shrink: 0; scroll-snap-align: start; }
  .modal-tool-btn:hover { background: var(--green-50); }
  .modal-tool-btn .ic { font-size: 22px; line-height: 1; }
  .modal-tool-btn.primary { background: var(--green-700); color: #fff; }
  .modal-tool-btn.primary:hover { background: var(--green-900); }
  .pdf-page-wrap { position: relative; margin: 46px auto 40px; }
  .pdf-page { position: relative; background: #fff; box-shadow: 0 4px 12px rgba(0,0,0,.25); transform-origin: top left; will-change: transform; }
  .pdf-page canvas { display: block; }
  .page-zoom-controls { position: absolute; top: -34px; left: 0; right: 0; display: flex; justify-content: center; gap: 6px; align-items: center; z-index: 5; }
  .page-zoom-controls .page-num { color: var(--gold-light); font-size: 12px; font-weight: 700; background: rgba(0,0,0,0.5); padding: 4px 10px; border-radius: 12px; }
  .page-zoom-controls button { background: rgba(255,255,255,0.9); color: var(--green-900); border: 1px solid rgba(0,0,0,0.1); padding: 4px 10px; font-size: 13px; border-radius: 12px; cursor: pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.2); }
  .page-zoom-controls .page-zoom-badge { background: rgba(255,255,255,0.9); padding: 4px 8px; border-radius: 12px; font-size: 11px; font-weight: 700; color: var(--green-900); min-width: 40px; text-align: center; }
  .page-num-btn { background: var(--gold) !important; color: #1a1a1a !important; font-weight: 700 !important; padding: 4px 12px !important; border-radius: 12px !important; }
  .page-num-btn:hover { background: #b58f1c !important; }
  /* Focus mode - hide other pages */
  .focus-mode .pdf-page-wrap { display: none; }
  .focus-mode .pdf-page-wrap.focused { display: block; margin: 46px auto 40px; }
  .focus-banner { position: sticky; top: 0; background: linear-gradient(90deg, var(--gold), #b58f1c); color: #1a1a1a; padding: 10px 16px; font-weight: 700; display: flex; justify-content: space-between; align-items: center; gap: 10px; flex-wrap: wrap; z-index: 30; box-shadow: 0 3px 10px rgba(0,0,0,0.3); border-radius: 8px; margin-bottom: 4px; width: 100%; max-width: 900px; }
  .focus-banner button { background: rgba(255,255,255,0.5); color: #1a1a1a; border: 1px solid rgba(0,0,0,0.2); padding: 6px 14px; font-weight: 700; }
  .focus-banner button:hover { background: rgba(255,255,255,0.8); }
  .page-navigator { background: rgba(255,255,255,0.95); border-radius: 10px; padding: 12px 16px; margin-bottom: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); max-width: 90vw; }
  .page-nav-label { font-size: 13px; color: var(--green-900); font-weight: 700; margin-bottom: 10px; }
  .page-nav-buttons { display: flex; gap: 6px; flex-wrap: wrap; }
  .page-nav-buttons button { background: var(--green-100); color: var(--green-900); font-weight: 700; padding: 6px 12px; font-size: 12px; border-radius: 8px; border: 1px solid var(--green-500); }
  .page-nav-buttons button:hover { background: var(--green-700); color: #fff; }
  .focus-mode .page-navigator { display: none; }
  /* Toast */
  .pgds-toast { position: fixed; bottom: 100px; left: 50%; transform: translateX(-50%); background: var(--green-900); color: #fff; padding: 12px 22px; border-radius: 24px; font-size: 14px; font-weight: 600; z-index: 3000; box-shadow: 0 6px 20px rgba(0,0,0,0.3); max-width: 90vw; text-align: center; animation: toastIn .25s ease; border: 1px solid var(--gold); }
  .pgds-toast.error { background: #991b1b; border-color: #fca5a5; }
  .pgds-toast.warn { background: #92400e; border-color: #fcd34d; }
  .pgds-toast.hide { opacity: 0; transition: opacity .4s; }
  @keyframes toastIn { from { opacity: 0; transform: translateX(-50%) translateY(12px); } to { opacity: 1; transform: translateX(-50%) translateY(0); } }
  .sig-overlay { position: absolute; border: 2px dashed var(--green-700); cursor: move; background: rgba(13,92,47,.04); user-select: none; touch-action: none; }
  .sig-overlay.locked { border: none; background: transparent; cursor: default; pointer-events: none; }
  .sig-overlay img, .sig-overlay .sig-text { width: 100%; height: 100%; pointer-events: none; display: block; }
  .sig-overlay .sig-text { display: flex; align-items: center; justify-content: center; font-family: 'Brush Script MT', cursive; font-size: 32px; color: var(--green-900); white-space: nowrap; }
  .sig-overlay .sig-text.stamp { font-family: Arial, sans-serif; font-weight: 900; color: var(--red); border: 3px solid var(--red); border-radius: 6px; }
  .sig-overlay .resize-handle { position: absolute; right: -8px; bottom: -8px; width: 20px; height: 20px; background: var(--green-700); border-radius: 10px; cursor: nwse-resize; }
  .sig-overlay .remove-handle { position: absolute; left: -8px; top: -8px; width: 24px; height: 24px; background: var(--red); color: #fff; border-radius: 12px; text-align: center; line-height: 24px; cursor: pointer; font-size: 14px; }
  /* Signature Pad fullscreen */
  .sigpad { position: fixed; inset: 0; background: #f8fafc; z-index: 2000; display: flex; flex-direction: column; }
  .sigpad-header { background: linear-gradient(135deg, #2563eb, #1e40af); color: #fff; padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
  .sigpad-header .title { font-size: 18px; font-weight: 700; display: flex; align-items: center; gap: 12px; }
  .sigpad-header .back { background: transparent; border: none; color: #fff; font-size: 24px; cursor: pointer; padding: 4px 10px; }
  .sigpad-header .tools { display: flex; gap: 10px; align-items: center; }
  .sigpad-header button { background: rgba(255,255,255,0.2); color: #fff; border: 1px solid rgba(255,255,255,0.3); padding: 8px 14px; border-radius: 20px; font-size: 13px; font-weight: 600; cursor: pointer; }
  .sigpad-header button:hover { background: rgba(255,255,255,0.3); }
  .sigpad-header button.save { background: #fff; color: #1e40af; }
  .sigpad-header button.trash { background: rgba(220,38,38,0.9); border-color: transparent; padding: 8px 12px; }
  .sigpad-canvas-wrap { flex: 1; padding: 20px; display: flex; align-items: center; justify-content: center; }
  .sigpad-canvas { background: #fff; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); width: 100%; height: 100%; max-height: 70vh; touch-action: none; cursor: crosshair; display: block; }
  .sigpad-controls { background: #fff; padding: 12px 20px; display: flex; justify-content: center; gap: 18px; align-items: center; box-shadow: 0 -2px 8px rgba(0,0,0,0.05); flex-wrap: wrap; }
  .color-swatches { display: flex; gap: 8px; }
  .color-swatch { width: 34px; height: 34px; border-radius: 17px; border: 3px solid transparent; cursor: pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.15); }
  .color-swatch.active { border-color: #64748b; transform: scale(1.1); }
  .pen-thickness { display: flex; gap: 6px; align-items: center; }
  .pen-btn { background: #f1f5f9; border: 2px solid transparent; padding: 6px 12px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 6px; font-size: 12px; font-weight: 600; color: var(--muted); }
  .pen-btn.active { background: #dbeafe; border-color: #2563eb; color: #1e40af; }
  .pen-btn .dot { display: inline-block; background: currentColor; border-radius: 50%; }
  .draw-canvas { border: 1px solid var(--border); border-radius: 8px; background: #fff; touch-action: none; cursor: crosshair; }
  .saved-sigs { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .saved-sig { border: 1px solid var(--border); border-radius: 6px; padding: 8px; cursor: pointer; text-align: center; position: relative; background: #fff; }
  .saved-sig:hover { border-color: var(--green-700); }
  .saved-sig img { max-width: 100%; max-height: 40px; }
  .saved-sig .del { position: absolute; top: 2px; right: 4px; color: var(--red); font-size: 11px; cursor: pointer; font-weight: 700; }
  /* Popover for saved / text / stamp picks from bottom bar */
  .popover { position: fixed; bottom: 90px; left: 50%; transform: translateX(-50%); background: #fff; border-radius: 12px; padding: 18px; box-shadow: 0 8px 24px rgba(0,0,0,0.2); z-index: 1100; max-width: 92vw; width: 380px; max-height: 60vh; overflow-y: auto; }
  .popover h4 { font-size: 14px; margin-bottom: 10px; color: var(--green-900); }
  .popover-close { position: absolute; top: 8px; right: 10px; background: transparent; border: none; font-size: 18px; cursor: pointer; color: var(--muted); }
  .popover-back { position: fixed; inset: 0; z-index: 1099; background: rgba(0,0,0,0.15); }
  /* Chart */
  .chart-box { background: var(--card); border-radius: 12px; padding: 16px; box-shadow: 0 2px 8px rgba(6,78,59,0.05); margin-bottom: 20px; }
  .chart-box h3 { font-size: 14px; color: var(--green-900); margin-bottom: 12px; font-weight: 700; }
  .bar-chart { display: flex; gap: 8px; align-items: flex-end; height: 160px; padding: 0 4px; overflow-x: auto; }
  .bar { flex: 1; min-width: 32px; background: linear-gradient(180deg, var(--green-500), var(--green-700)); border-radius: 4px 4px 0 0; position: relative; }
  .bar .lbl { position: absolute; bottom: -22px; left: 0; right: 0; text-align: center; font-size: 10px; color: var(--muted); }
  .bar .val { position: absolute; top: -18px; left: 0; right: 0; text-align: center; font-size: 11px; font-weight: 700; color: var(--green-900); }
  /* Install banner */
  .install-banner { position: fixed; bottom: 16px; left: 16px; right: 16px; max-width: 440px; margin: 0 auto; background: var(--green-700); color: #fff; padding: 12px 16px; border-radius: 12px; display: flex; gap: 12px; align-items: center; z-index: 500; box-shadow: 0 8px 24px rgba(0,0,0,0.25); }
  .install-banner img { width: 40px; height: 40px; border-radius: 20px; }
  .install-banner .txt { flex: 1; font-size: 13px; }
  .install-banner button { background: var(--gold); color: #1a1a1a; padding: 6px 14px; font-size: 13px; }
  .install-banner .close-x { background: transparent; padding: 0 6px; }
  /* Mobile */
  @media (max-width: 720px) {
    .container { padding: 10px; }
    .row2 { grid-template-columns: 1fr; }
    .header { padding: 10px 12px; }
    .brand-text .title { font-size: 15px; }
    .brand-text .subtitle { font-size: 10px; }
    .user-info { font-size: 12px; }
    th, td { padding: 8px 10px; font-size: 12px; }
    .card { padding: 16px; }
    .modal-tool-btn { min-width: 56px; padding: 6px 8px; font-size: 10px; }
    .modal-tool-btn .ic { font-size: 18px; }
    .modal-doc { padding: 12px 8px 100px; }
    .sigpad-header { padding: 10px 12px; }
    .sigpad-controls { gap: 10px; padding: 10px; }
    .color-swatch { width: 30px; height: 30px; }
  }
</style>
</head>
<body>

<div id="app"></div>
<div id="installBanner"></div>

<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script>
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
const SUPABASE_URL = 'https://icljtvutpjhxlnygxqbg.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImljbGp0dnV0cGpoeGxueWd4cWJnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODY2ODUxNzEsImV4cCI6MjEwMjI2MTE3MX0.c4VKSusY-sukSEzy6mD_lDnyXpad30idGjo5yohvWng';
const MAX_FILE_MB = 5;
const APP_VERSION = 'v10-multistore';
console.log('PG Docs Sign', APP_VERSION);

// Register service worker for PWA
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('/sw.js').catch(e => console.warn('SW reg failed', e));
}
let deferredInstallPrompt = null;
window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredInstallPrompt = e;
  showInstallBanner();
});
function showInstallBanner() {
  if (!deferredInstallPrompt) return;
  document.getElementById('installBanner').innerHTML = `
    <div class="install-banner">
      <img src="/SignDocs_icon.png">
      <div class="txt"><strong>Install PG Docs Sign</strong><br><span style="opacity:.85;">Add to home screen for quick access</span></div>
      <button onclick="doInstall()">Install</button>
      <button class="close-x" onclick="dismissInstall()">&times;</button>
    </div>`;
}
function doInstall() {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  deferredInstallPrompt.userChoice.finally(() => { deferredInstallPrompt = null; dismissInstall(); });
}
function dismissInstall() { document.getElementById('installBanner').innerHTML = ''; }

const { createClient } = supabase;
const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const app = document.getElementById('app');
let currentUser = null;
let authMode = 'login';
let adminView = 'dashboard';
let docsFilter = 'pending';
let staffTab = 'pending';
let staffView = 'docs';
let selectedFiles = [];

function h(s) { return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function showMsg(id, text, type = 'info') {
  const el = document.getElementById(id);
  if (el) { el.innerHTML = `<div class="msg ${type}">${text}</div>`; setTimeout(() => el.innerHTML = '', 5000); }
}
function ageHours(from, to) {
  return (new Date(to || Date.now()) - new Date(from)) / 36e5;
}
function ageBadge(hours) {
  const h = Math.floor(hours);
  let cls = 'fresh', label = h + 'h';
  if (hours >= 24) { cls = 'warm'; label = Math.floor(hours/24) + 'd ' + (h%24) + 'h'; }
  if (hours >= 48) cls = 'hot';
  if (hours >= 72) cls = 'critical';
  return `<span class="age-badge ${cls}">${label}</span>`;
}

// ---------- ROUTER ----------
async function route() {
  const { data: { session } } = await sb.auth.getSession();
  if (!session) { authMode = 'login'; return renderAuth(); }
  const { data: profile } = await sb.from('pgds_profiles').select('*').eq('id', session.user.id).single();
  if (!profile) { await sb.auth.signOut(); return renderAuth(); }
  if (profile.status !== 'approved') {
    await sb.auth.signOut(); authMode = 'login'; renderAuth();
    showMsg('authMsg', profile.status === 'rejected' ? 'Your account was not approved. Reason: ' + (profile.reject_reason || 'not specified') : 'Your account is still awaiting admin approval.', profile.status === 'rejected' ? 'error' : 'info');
    return;
  }
  currentUser = { id: session.user.id, email: session.user.email, ...profile };
  // Load every store this user is allowed to upload for
  currentUser.stores = await loadMyStores(session.user.id, profile);
  if (profile.role === 'admin') renderAdmin(); else renderStaff();
}

async function loadMyStores(userId, profile) {
  const { data } = await sb.from('pgds_user_stores')
    .select('store_id, pgds_stores(id, store_number, store_name, active)')
    .eq('user_id', userId);
  const list = (data || [])
    .map(r => r.pgds_stores)
    .filter(s => s && s.active !== false)
    .sort((a, b) => String(a.store_number).localeCompare(String(b.store_number), undefined, { numeric: true }));
  // Fall back to the profile's single store if no assignments exist yet
  if (!list.length && profile.store_number) {
    return [{ id: null, store_number: profile.store_number, store_name: profile.store_name || '' }];
  }
  return list;
}

// ---------- AUTH ----------
async function renderAuth() {
  const stores = await sb.from('pgds_stores').select('*').eq('active', true).order('store_number', { ascending: true });
  const storeOptions = (stores.data || []).map(s => `<option value="${s.id}" data-num="${h(s.store_number)}" data-name="${h(s.store_name)}">${h(s.store_number)} - ${h(s.store_name)}</option>`).join('');
  app.innerHTML = `
    <div class="center">
      <div class="card login-card">
        <div class="login-header">
          <img src="/SignDocs_icon.png" alt="PG Docs Sign">
          <h1>PG Docs Sign</h1>
          <div class="tag">Digital Signature Portal</div>
        </div>
        <div id="authMsg"></div>
        <form id="authForm">
          ${authMode === 'signup' ? `
            <div class="form-group"><label>Full name</label><input type="text" id="fullName" required placeholder="e.g. Juan Dela Cruz"></div>
            <div class="form-group">
              <label>Select your store(s) *</label>
              <p style="font-size:12px;color:var(--muted);margin:-2px 0 8px;">Tick every store you handle. You can pick more than one.</p>
              <div class="store-picker" id="storePicker">
                ${(stores.data || []).map(s => `
                  <label class="store-opt">
                    <input type="checkbox" value="${s.id}" data-num="${h(s.store_number)}" data-name="${h(s.store_name)}">
                    <span>${h(s.store_number)} - ${h(s.store_name)}</span>
                  </label>`).join('') || '<p style="font-size:13px;color:var(--muted);">No stores configured yet - ask the admin.</p>'}
              </div>
            </div>
          ` : ''}
          <div class="form-group"><label>Email</label><input type="email" id="email" required placeholder="you@example.com"></div>
          <div class="form-group"><label>Password</label><input type="password" id="password" required minlength="6" placeholder="At least 6 characters"></div>
          <button type="submit" class="btn-full" id="submitBtn">${authMode === 'signup' ? 'Create account' : 'Sign in'}</button>
        </form>
        <p style="text-align:center;margin-top:16px;font-size:13px;color:var(--muted);">
          ${authMode === 'signup' ? 'Already have an account?' : 'New here?'}
          <span class="link" id="switchLink">${authMode === 'signup' ? 'Sign in' : 'Create an account'}</span>
        </p>
      </div>
    </div>`;
  document.getElementById('switchLink').onclick = () => { authMode = authMode === 'login' ? 'signup' : 'login'; renderAuth(); };
  document.getElementById('authForm').onsubmit = handleAuth;
}

async function handleAuth(e) {
  e.preventDefault();
  const email = document.getElementById('email').value.trim();
  const password = document.getElementById('password').value;
  const btn = document.getElementById('submitBtn'); btn.disabled = true;
  try {
    if (authMode === 'signup') {
      const fullName = document.getElementById('fullName').value.trim();
      const picked = Array.from(document.querySelectorAll('#storePicker input:checked'));
      if (!picked.length) throw new Error('Please tick at least one store.');
      const { data, error } = await sb.auth.signUp({ email, password, options: { data: { full_name: fullName } } });
      if (error) throw error;
      if (data.user) {
        // Primary store = the first one ticked (used for display)
        await sb.from('pgds_profiles').update({
          full_name: fullName,
          store_number: picked[0].dataset.num,
          store_name: picked[0].dataset.name
        }).eq('id', data.user.id);
        await sb.from('pgds_user_stores').insert(
          picked.map(p => ({ user_id: data.user.id, store_id: p.value }))
        );
      }
      showMsg('authMsg', `Account created for ${picked.length} store(s)! Waiting for admin approval.`, 'success');
    } else {
      const { data, error } = await sb.auth.signInWithPassword({ email, password });
      if (error) throw error;
      await route();
    }
  } catch (err) { showMsg('authMsg', err.message || 'Error', 'error'); }
  finally { btn.disabled = false; }
}

// ---------- STAFF ----------
// ---------- GUIDE ----------
function guideHTML() {
  const rules = [
    { n: 1, icon: '✍️', title: 'Complete the signatories first',
      body: 'All documents must already carry the required handwritten signatures - Store Manager, Area Manager, and any other concerned personnel - <strong>before</strong> you upload them here. Skip this only when the document genuinely does not require those signatures.' },
    { n: 2, icon: '📑', title: 'One document title per upload',
      body: 'Do not mix different document types in a single submission. Every file in one upload must belong to the <strong>same document title or category</strong>. Mixed submissions may be rejected.' },
    { n: 3, icon: '🔍', title: 'Upload clear, readable copies',
      body: 'Photos and scans must be sharp, well lit, properly aligned, and complete. <strong>Blurry, dark, cropped, or unreadable</strong> files may be rejected.' },
    { n: 4, icon: '📱', title: 'Phone or computer both work',
      body: 'You can use this app on a mobile phone or on a PC. All that is required is a <strong>stable internet connection</strong>. On a phone you may upload from your gallery or take a photo directly.' },
    { n: 5, icon: '🚦', title: 'Use the urgency level honestly',
      body: 'Choose the urgency that truly matches the document. Reserve <strong>Urgent</strong> and <strong>High</strong> for items that genuinely need immediate attention.' },
    { n: 6, icon: '🗓️', title: 'Respect the signing schedule',
      body: 'Documents are reviewed and signed on an organized schedule so signing time is used efficiently. Honest use of the urgency option ensures that genuinely urgent documents get prioritized.' },
  ];
  return `
    <div class="guide-hero">
      <img src="/SignDocs_icon.png" alt="">
      <div>
        <h2>Admin Staff Guide</h2>
        <p>Document uploading standards - please read before submitting</p>
      </div>
    </div>

    <div class="guide-grid">
      ${rules.map(r => `
        <div class="guide-card">
          <div class="guide-card-head">
            <span class="guide-num">${r.n}</span>
            <span class="guide-icon">${r.icon}</span>
            <span class="guide-title">${r.title}</span>
          </div>
          <p class="guide-body">${r.body}</p>
        </div>`).join('')}
    </div>

    <div class="guide-compare">
      <div class="guide-good">
        <h4>✅ Accepted</h4>
        <ul>
          <li>Signed by all required personnel</li>
          <li>One document title per upload</li>
          <li>Sharp, bright, fully visible page</li>
          <li>Whole document in frame, edges included</li>
          <li>Urgency matches the real need</li>
        </ul>
      </div>
      <div class="guide-bad">
        <h4>❌ Likely rejected</h4>
        <ul>
          <li>Missing Store or Area Manager signature</li>
          <li>Several different document types mixed together</li>
          <li>Blurry, dark, or glare-covered photo</li>
          <li>Corners or lines cut off</li>
          <li>Everything marked Urgent</li>
        </ul>
      </div>
    </div>

    <div class="guide-steps">
      <h4>📤 How to upload</h4>
      <ol>
        <li><strong>Collect the signed document</strong> - make sure every required signature is already on paper.</li>
        <li><strong>Give it a clear title</strong> - e.g. "RFP - Danlhey Paint Center - Sept 2026".</li>
        <li><strong>Choose the urgency</strong> that honestly reflects the deadline.</li>
        <li><strong>Add the files</strong> - tap <em>Choose files</em> to pick from your device, or <em>Take photo</em> to capture pages with your camera. Multiple pages of the <em>same</em> document are merged into one file automatically.</li>
        <li><strong>Add notes</strong> if there is anything the signer should know.</li>
        <li><strong>Tap Upload for signature</strong> and watch the status on your dashboard.</li>
      </ol>
    </div>

    <div class="guide-status">
      <h4>📊 What the statuses mean</h4>
      <div class="guide-status-row"><span class="badge pending">Pending</span><span>Waiting in the signer's queue. You may still cancel it.</span></div>
      <div class="guide-status-row"><span class="badge approved">Approved</span><span>Signed and finalized. Tap Download to get the signed copy.</span></div>
      <div class="guide-status-row"><span class="badge rejected">Rejected</span><span>Returned with a reason. Fix the issue and upload again.</span></div>
      <div class="guide-status-row"><span style="font-size:18px;">📣</span><span>A gold note on your document is a <strong>message from the signer</strong> - usually a missing requirement. Read it and act on it.</span></div>
    </div>

    <div class="guide-footer">
      Before submitting, please review your documents carefully to avoid rejection and unnecessary delays.
    </div>`;
}

function niceName() {
  if (currentUser.full_name && currentUser.full_name.trim()) return currentUser.full_name;
  const local = (currentUser.email || '').split('@')[0].replace(/[._-]/g, ' ');
  return local.split(' ').filter(Boolean).map(w => w[0].toUpperCase() + w.slice(1)).join(' ') || 'User';
}
function headerHTML(subtitle) {
  const st = currentUser.stores || [];
  let storeLabel = '';
  if (st.length > 1) storeLabel = `${st.length} stores: ` + st.map(s => s.store_number).join(', ');
  else if (st.length === 1) storeLabel = `Store #${st[0].store_number} - ${st[0].store_name}`;
  else storeLabel = [currentUser.store_number && ('Store #' + currentUser.store_number), currentUser.store_name].filter(Boolean).join(' - ') || '';
  const uname = niceName();
  const personSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="#f7e08a" style="filter:drop-shadow(0 1px 1px rgba(0,0,0,0.3));"><path d="M12 12c2.7 0 4.9-2.2 4.9-4.9S14.7 2.2 12 2.2 7.1 4.4 7.1 7.1 9.3 12 12 12zm0 2.4c-3.3 0-9.8 1.6-9.8 4.9v2.5h19.6v-2.5c0-3.3-6.5-4.9-9.8-4.9z"/></svg>`;
  return `<div class="header">
    <div class="header-row1">
      <div class="brand">
        <img src="/SignDocs_icon.png" alt="">
        <div class="brand-text">
          <span class="title">PG Docs Sign</span>
          <span class="subtitle">${subtitle}</span>
          <div class="header-user" style="margin-top:4px;">
            <span class="uname" onclick="editMyName()">${personSvg} ${h(uname)} <span class="edit-icon">✏️</span></span>
            ${storeLabel ? `<span class="ustore" style="margin-left:26px;">${h(storeLabel)}</span>` : ''}
          </div>
        </div>
      </div>
      <button class="menu-btn" onclick="toggleMenu(event)" title="Menu">⋮</button>
    </div>
    <div id="headerMenu" style="display:none;" class="menu-drop">
      <button onclick="editMyName()">✏️ Edit my name</button>
      <hr>
      <button onclick="refreshApp()">🔄 Refresh</button>
      <hr>
      <button onclick="logout()">🚪 Logout</button>
    </div>
  </div>`;
}

async function editMyName() {
  const current = currentUser.full_name || '';
  const newName = prompt('Enter your full name:', current);
  if (newName === null) return;
  const trimmed = newName.trim();
  if (!trimmed) return;
  const { error } = await sb.from('pgds_profiles').update({ full_name: trimmed }).eq('id', currentUser.id);
  if (error) return alert(error.message);
  currentUser.full_name = trimmed;
  if (currentUser.role === 'admin') renderAdmin(); else renderStaff();
}
function toggleMenu(e) {
  e.stopPropagation();
  const m = document.getElementById('headerMenu');
  if (!m) return;
  m.style.display = m.style.display === 'none' ? 'block' : 'none';
  const close = (ev) => { if (!m.contains(ev.target) && !ev.target.classList.contains('menu-btn')) { m.style.display = 'none'; document.removeEventListener('click', close); } };
  setTimeout(() => document.addEventListener('click', close), 10);
}

async function renderStaff() {
  if (staffView === 'guide') {
    app.innerHTML = `
      ${headerHTML('Guide')}
      <div class="container">
        <div class="tabs">
          <div class="tab" data-sv="docs">📄 My Documents</div>
          <div class="tab active" data-sv="guide">📖 Admin Staff Guide</div>
        </div>
        ${guideHTML()}
      </div>`;
    document.querySelectorAll('.tab[data-sv]').forEach(t => t.onclick = () => { staffView = t.dataset.sv; renderStaff(); });
    return;
  }
  app.innerHTML = `
    ${headerHTML('My Documents')}
    <div class="container">
      <div class="tabs">
        <div class="tab active" data-sv="docs">📄 My Documents</div>
        <div class="tab" data-sv="guide">📖 Admin Staff Guide</div>
      </div>
      <div id="staffKpis" class="kpi-row"></div>
      <div class="box">
        <h3>Upload documents for signature</h3>
        <div class="upload-note">
          <strong>Note:</strong> Please file / upload documents <strong>per title or per document name</strong>. If a single document has multiple pages or photos, add them all together here — they will be merged into one signed file.
        </div>
        <div id="uploadMsg"></div>
        <form id="uploadForm">
          ${(currentUser.stores || []).length > 1 ? `
          <div class="form-group">
            <label>Which store is this document for? *</label>
            <select id="docStore" required>
              ${currentUser.stores.map((s, i) => `<option value="${i}">${h(s.store_number)} - ${h(s.store_name)}</option>`).join('')}
            </select>
          </div>` : ''}
          <div class="form-group"><label>Document title *</label><input type="text" id="docTitle" required placeholder="e.g. September expense report"></div>
          <div class="row2">
            <div class="form-group"><label>Urgency</label><select id="docUrgency">
              <option value="low">Low</option><option value="normal" selected>Normal</option>
              <option value="high">High</option><option value="urgent">Urgent</option>
            </select></div>
            <div class="form-group"><label>Files (PDF/JPG/PNG, max 5MB each) *</label>
              <label for="fileInput" class="file-drop" id="fileDrop"><span id="fileDropText">Choose files or take a photo</span></label>
              <input type="file" id="fileInput" accept=".pdf,.jpg,.jpeg,.png" multiple>
              <input type="file" id="cameraInput" accept="image/*" capture="environment" multiple>
              <div class="upload-buttons">
                <button type="button" class="secondary" onclick="document.getElementById('fileInput').click()">📁 Choose files</button>
                <button type="button" class="gold" onclick="document.getElementById('cameraInput').click()">📷 Take photo</button>
              </div>
              <div id="fileList" class="file-list"></div>
            </div>
          </div>
          <div class="form-group"><label>Notes (optional)</label><textarea id="docNotes" placeholder="Any note for the signer..."></textarea></div>
          <button type="submit" id="uploadBtn">Upload for signature</button>
        </form>
      </div>
      <h2 style="font-size:16px;margin-bottom:12px;color:var(--green-900);">My Documents</h2>
      <div class="tabs">
        <div class="tab ${staffTab==='pending'?'active':''}" data-t="pending">Pending</div>
        <div class="tab ${staffTab==='approved'?'active':''}" data-t="approved">Approved</div>
        <div class="tab ${staffTab==='rejected'?'active':''}" data-t="rejected">Rejected</div>
      </div>
      <div id="docList"></div>
    </div>`;
  document.querySelectorAll('.tab[data-sv]').forEach(t => t.onclick = () => { staffView = t.dataset.sv; renderStaff(); });
  document.querySelectorAll('.tab[data-t]').forEach(t => t.onclick = () => { staffTab = t.dataset.t; renderStaff(); });
  document.getElementById('fileInput').onchange = handleFilePick;
  document.getElementById('cameraInput').onchange = handleFilePick;
  document.getElementById('uploadForm').onsubmit = handleUpload;
  selectedFiles = [];
  updateFileList();
  loadStaffKpis();
  loadMyDocs();
}

async function loadStaffKpis() {
  const el = document.getElementById('staffKpis');
  const { data } = await sb.from('pgds_documents').select('status,created_at,signed_at').eq('uploader_id', currentUser.id);
  const pending = (data||[]).filter(d => d.status === 'pending');
  const approved = (data||[]).filter(d => d.status === 'approved');
  const oldestPending = pending.reduce((min, d) => Math.max(min, ageHours(d.created_at)), 0);
  const avgTurnaround = approved.length ? approved.reduce((s, d) => s + ageHours(d.created_at, d.signed_at), 0) / approved.length : 0;
  el.innerHTML = `
    <div class="kpi gold"><div class="label">My Pending</div><div class="value">${pending.length}</div><div class="sub">${oldestPending ? 'Oldest: ' + Math.floor(oldestPending) + 'h' : 'All caught up'}</div></div>
    <div class="kpi"><div class="label">Approved</div><div class="value">${approved.length}</div><div class="sub">Total signed</div></div>
    <div class="kpi blue"><div class="label">Avg turnaround</div><div class="value">${avgTurnaround.toFixed(1)}h</div><div class="sub">From upload to signing</div></div>`;
}

function handleFilePick(e) {
  const files = Array.from(e.target.files);
  for (const f of files) {
    const mb = f.size / 1048576;
    if (mb > MAX_FILE_MB) { alert(`${f.name} is too big (${mb.toFixed(1)}MB). Max ${MAX_FILE_MB}MB per file.`); continue; }
    selectedFiles.push(f);
  }
  e.target.value = '';
  updateFileList();
}
function updateFileList() {
  const list = document.getElementById('fileList');
  const drop = document.getElementById('fileDrop');
  const txt = document.getElementById('fileDropText');
  if (!selectedFiles.length) {
    list.innerHTML = ''; drop.classList.remove('has-file'); txt.textContent = 'Choose files or take a photo';
    return;
  }
  drop.classList.add('has-file');
  txt.textContent = `${selectedFiles.length} file(s) selected - they will be merged`;
  list.innerHTML = selectedFiles.map((f, i) => `<div class="file-item"><span>📄 ${h(f.name)} (${(f.size/1048576).toFixed(2)}MB)</span><span class="rm" onclick="removeFile(${i})">Remove</span></div>`).join('');
}
function removeFile(i) { selectedFiles.splice(i, 1); updateFileList(); }

// Downscale + re-encode a photo so the resulting PDF stays small and fast to re-save.
// Documents stay perfectly readable at 2000px on the long edge.
async function compressImage(file, maxDim = 2000, quality = 0.82) {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, maxDim / Math.max(bitmap.width, bitmap.height));
  const w = Math.round(bitmap.width * scale), hgt = Math.round(bitmap.height * scale);
  const canvas = document.createElement('canvas');
  canvas.width = w; canvas.height = hgt;
  const ctx = canvas.getContext('2d');
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(bitmap, 0, 0, w, hgt);
  bitmap.close && bitmap.close();
  const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', quality));
  return { bytes: new Uint8Array(await blob.arrayBuffer()), width: w, height: hgt };
}

async function mergeFilesToPdf(files) {
  const pdfDoc = await PDFLib.PDFDocument.create();
  for (const f of files) {
    const ext = (f.name.split('.').pop() || '').toLowerCase();
    if (ext === 'pdf') {
      const buf = new Uint8Array(await f.arrayBuffer());
      const src = await PDFLib.PDFDocument.load(buf);
      const pages = await pdfDoc.copyPages(src, src.getPageIndices());
      pages.forEach(p => pdfDoc.addPage(p));
    } else {
      let img, iw, ih;
      try {
        const c = await compressImage(f);
        img = await pdfDoc.embedJpg(c.bytes); iw = c.width; ih = c.height;
      } catch (e) {
        // Fallback to the original bytes if canvas compression is unavailable
        const buf = new Uint8Array(await f.arrayBuffer());
        img = (['jpg','jpeg'].includes(ext) || f.type === 'image/jpeg') ? await pdfDoc.embedJpg(buf) : await pdfDoc.embedPng(buf);
        iw = img.width; ih = img.height;
      }
      const page = pdfDoc.addPage([iw, ih]);
      page.drawImage(img, { x: 0, y: 0, width: iw, height: ih });
    }
  }
  return await pdfDoc.save({ useObjectStreams: false });
}

async function handleUpload(e) {
  e.preventDefault();
  if (!selectedFiles.length) return alert('Please choose at least one file.');
  const title = document.getElementById('docTitle').value.trim();
  const urgency = document.getElementById('docUrgency').value;
  const notes = document.getElementById('docNotes').value.trim();
  if (!title) return;
  const btn = document.getElementById('uploadBtn'); btn.disabled = true; btn.textContent = 'Uploading...';
  try {
    const pdfBytes = selectedFiles.length === 1 && selectedFiles[0].name.toLowerCase().endsWith('.pdf') ? new Uint8Array(await selectedFiles[0].arrayBuffer()) : await mergeFilesToPdf(selectedFiles);
    const path = `${currentUser.id}/${Date.now()}-${Math.random().toString(36).slice(2,8)}.pdf`;
    const up = await sb.storage.from('pgds-pending').upload(path, pdfBytes, { contentType: 'application/pdf' });
    if (up.error) throw up.error;
    // Which store did they pick? (single-store users just use their only one)
    const stores = currentUser.stores || [];
    const sel = document.getElementById('docStore');
    const chosen = stores.length > 1 && sel ? stores[Number(sel.value)] : (stores[0] || null);
    const ins = await sb.from('pgds_documents').insert({
      uploader_id: currentUser.id,
      store_number: chosen ? chosen.store_number : currentUser.store_number,
      store_name: chosen ? chosen.store_name : currentUser.store_name,
      title, notes, urgency, original_file_path: path, status: 'pending'
    });
    if (ins.error) { await sb.storage.from('pgds-pending').remove([path]); throw ins.error; }
    showMsg('uploadMsg', `Uploaded! ${selectedFiles.length} file(s) merged into one signed queue.`, 'success');
    document.getElementById('uploadForm').reset();
    selectedFiles = []; updateFileList();
    staffTab = 'pending'; renderStaff();
  } catch (err) { showMsg('uploadMsg', err.message || err, 'error'); }
  finally { btn.disabled = false; btn.textContent = 'Upload for signature'; }
}

async function loadMyDocs() {
  const list = document.getElementById('docList');
  list.innerHTML = '<p class="empty">Loading...</p>';
  const { data, error } = await sb.from('pgds_documents').select('*').eq('uploader_id', currentUser.id).eq('status', staffTab).order('created_at', { ascending: false });
  if (error) return list.innerHTML = `<div class="msg error">${error.message}</div>`;
  if (!data.length) return list.innerHTML = `<p class="empty">No ${staffTab} documents.</p>`;
  const multi = (currentUser.stores || []).length > 1;
  list.innerHTML = `<table><thead><tr><th>Title</th>${multi ? '<th>Store</th>' : ''}<th>Urgency</th><th>Uploaded</th><th>Age</th><th>Status</th><th>Actions</th></tr></thead><tbody>
    ${data.map(d => { const age = ageHours(d.created_at, d.signed_at); return `<tr>
      <td><strong>${h(d.title)}</strong>${d.notes ? `<div style="font-size:12px;color:#64748b;margin-top:4px;">${h(d.notes)}</div>` : ''}${d.admin_remarks ? `<div class="staff-remark">📣 <strong>Message from signer:</strong><br>${h(d.admin_remarks)}</div>` : ''}</td>
      ${multi ? `<td>${h(d.store_number || '-')}${d.store_name ? `<div style="font-size:11px;color:var(--muted);">${h(d.store_name)}</div>` : ''}</td>` : ''}
      <td><span class="badge ${d.urgency}">${d.urgency}</span></td>
      <td>${new Date(d.created_at).toLocaleString()}</td>
      <td>${ageBadge(age)}${d.signed_at ? `<div style="font-size:11px;color:var(--muted);">until signed</div>` : ''}</td>
      <td><span class="badge ${d.status}">${d.status}</span>${d.status==='rejected'&&d.reject_reason?`<div style="font-size:12px;color:#991b1b;margin-top:4px;">${h(d.reject_reason)}</div>`:''}</td>
      <td class="actions">
        ${d.status==='approved'&&d.signed_file_path?`<button onclick="downloadDirect('pgds-signed','${d.signed_file_path}','${h(safeFileName(d.title)).replace(/'/g,"\\'")}')">⬇️ Download</button><button class="secondary" onclick="downloadFile('pgds-signed','${d.signed_file_path}')">👁️ View</button>`:''}
        ${d.status==='pending'?`<button class="secondary" onclick="downloadFile('pgds-pending','${d.original_file_path}')">👁️ Preview</button>`:''}
        ${d.status==='pending'?`<button class="danger" onclick="cancelDoc('${d.id}','${d.original_file_path}')">Cancel</button>`:''}
      </td>
    </tr>`; }).join('')}</tbody></table>`;
}

async function downloadFile(bucket, path) {
  await openInAppViewer(bucket, path);
}

// Make a safe, meaningful filename from a document title
function safeFileName(title, suffix = 'signed') {
  const base = String(title || 'document')
    .replace(/[^a-zA-Z0-9 _.-]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 70) || 'document';
  return `${base} - ${suffix}.pdf`;
}

// Download straight to the device's Downloads folder (phone or PC)
async function downloadDirect(bucket, path, filename) {
  toast('Preparing download...');
  const { data, error } = await sb.storage.from(bucket).createSignedUrl(path, 300, { download: filename });
  if (error) return toast(error.message, 'error');
  try {
    const resp = await fetch(data.signedUrl);
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 3000);
    toast('⬇️ Saved to your Downloads folder');
  } catch (e) {
    // Fallback: the signed URL already carries Content-Disposition: attachment
    window.location.href = data.signedUrl;
  }
}

// In-app PDF/image viewer modal (never leaves the PWA)
async function openInAppViewer(bucket, path) {
  const { data, error } = await sb.storage.from(bucket).createSignedUrl(path, 300);
  if (error) return alert(error.message);
  const url = data.signedUrl;
  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.id = 'viewerModal';
  const ext = (path.split('.').pop() || '').toLowerCase();
  const isImage = ['jpg','jpeg','png','gif','webp'].includes(ext);
  modal.innerHTML = `
    <div class="modal-topbar">
      <span class="title">👁️ ${h(path.split('/').pop())}</span>
      <div style="display:flex;gap:8px;">
        <button onclick="downloadDirect('${bucket}','${path}','${h(path.split('/').pop())}')">⬇️ Download</button>
        <button onclick="closeInAppViewer()">Close</button>
      </div>
    </div>
    <div class="modal-doc" id="viewerBody" style="padding:20px;">
      <p style="color:#fff">Loading...</p>
    </div>`;
  document.body.appendChild(modal);
  const body = document.getElementById('viewerBody');
  if (isImage) {
    body.innerHTML = `<img src="${url}" style="max-width:100%;height:auto;background:#fff;box-shadow:0 4px 12px rgba(0,0,0,.25);">`;
  } else {
    // Render PDF via PDF.js so it works reliably inside PWA
    try {
      const resp = await fetch(url);
      const buf = await resp.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
      body.innerHTML = '';
      const containerWidth = Math.max(320, body.clientWidth - 40);
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const raw = page.getViewport({ scale: 1 });
        const scale = Math.min(1.6, containerWidth / raw.width);
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.className = 'pdf-page';
        canvas.width = viewport.width; canvas.height = viewport.height;
        canvas.style.marginBottom = '12px';
        body.appendChild(canvas);
        await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      }
    } catch (err) {
      body.innerHTML = `<div class="msg error">Could not render: ${err.message}. <a href="${url}" target="_blank" style="color:#fff;">Open in new tab</a></div>`;
    }
  }
}

function closeInAppViewer() {
  const m = document.getElementById('viewerModal'); if (m) m.remove();
}

async function cancelDoc(id, path) {
  if (!confirm('Cancel this pending document? This cannot be undone.')) return;
  const { data: deleted, error } = await sb.from('pgds_documents').delete().eq('id', id).select('id');
  if (error) return toast(error.message, 'error');
  if (!deleted || !deleted.length) return toast('Could not cancel - permission denied.', 'error');
  try { await sb.storage.from('pgds-pending').remove([path]); } catch (e) {}
  toast('Document cancelled');
  loadMyDocs();
}

// ---------- ADMIN ----------
async function renderAdmin() {
  app.innerHTML = `
    ${headerHTML('Admin Console')}
    <div class="container">
      <div class="tabs">
        <div class="tab ${adminView==='dashboard'?'active':''}" data-v="dashboard">📊 Dashboard</div>
        <div class="tab ${adminView==='pending-users'?'active':''}" data-v="pending-users">Pending Users <span id="pendingCount"></span></div>
        <div class="tab ${adminView==='users'?'active':''}" data-v="users">All Users</div>
        <div class="tab ${adminView==='documents'?'active':''}" data-v="documents">📋 Signing Queue <span id="docsCount"></span></div>
        <div class="tab ${adminView==='approved-folders'?'active':''}" data-v="approved-folders">📁 Approved Docs</div>
        <div class="tab ${adminView==='rejected'?'active':''}" data-v="rejected">❌ Rejected</div>
        <div class="tab ${adminView==='stores'?'active':''}" data-v="stores">Stores</div>
        <div class="tab ${adminView==='signatures'?'active':''}" data-v="signatures">My Signatures</div>
        <div class="tab ${adminView==='guide'?'active':''}" data-v="guide">📖 Admin Staff Guide</div>
      </div>
      <div id="msg"></div>
      <div id="content"></div>
    </div>`;
  document.querySelectorAll('.tab[data-v]').forEach(t => t.onclick = () => { adminView = t.dataset.v; renderAdmin(); });
  renderAdminView();
  updateAdminCounts();
}

async function updateAdminCounts() {
  const [{ count: pu }, { count: pd }] = await Promise.all([
    sb.from('pgds_profiles').select('*', { count: 'exact', head: true }).eq('status', 'pending'),
    sb.from('pgds_documents').select('*', { count: 'exact', head: true }).eq('status', 'pending')
  ]);
  const pc = document.getElementById('pendingCount'); if (pc) pc.textContent = pu ? `(${pu})` : '';
  const dc = document.getElementById('docsCount'); if (dc) dc.textContent = pd ? `(${pd})` : '';
}

async function renderAdminView() {
  const el = document.getElementById('content');
  if (adminView === 'dashboard') return renderDashboard(el);
  if (adminView === 'pending-users') return renderPendingUsers(el);
  if (adminView === 'users') return renderAllUsers(el);
  if (adminView === 'documents') return renderPendingQueue(el);
  if (adminView === 'approved-folders') return renderApprovedFolders(el);
  if (adminView === 'rejected') return renderRejected(el);
  if (adminView === 'stores') return renderStores(el);
  if (adminView === 'signatures') return renderSignaturesMgr(el);
  if (adminView === 'guide') return el.innerHTML = guideHTML();
}

async function renderDashboard(el) {
  el.innerHTML = '<p class="empty">Loading dashboard...</p>';
  const { data: docs } = await sb.from('pgds_documents').select('id,status,created_at,signed_at,store_number,urgency');
  const all = docs || [];
  const pending = all.filter(d => d.status === 'pending');
  const approved = all.filter(d => d.status === 'approved');
  const rejected = all.filter(d => d.status === 'rejected');
  const today = new Date(); today.setHours(0,0,0,0);
  const signedToday = approved.filter(d => d.signed_at && new Date(d.signed_at) >= today).length;
  const oldestPending = pending.reduce((max, d) => Math.max(max, ageHours(d.created_at)), 0);
  const avgWait = approved.length ? approved.reduce((s, d) => s + ageHours(d.created_at, d.signed_at), 0) / approved.length : 0;
  const critical = pending.filter(d => ageHours(d.created_at) >= 48).length;
  const urgent = pending.filter(d => d.urgency === 'urgent' || d.urgency === 'high').length;
  // Docs per store
  const byStore = {};
  pending.forEach(d => { const k = d.store_number || 'unassigned'; byStore[k] = (byStore[k] || 0) + 1; });
  const storeEntries = Object.entries(byStore).sort((a,b) => b[1]-a[1]).slice(0, 10);
  const maxCount = Math.max(1, ...storeEntries.map(e => e[1]));

  el.innerHTML = `
    <div class="kpi-row">
      <div class="kpi gold"><div class="label">Pending Now</div><div class="value">${pending.length}</div><div class="sub">${urgent} urgent/high</div></div>
      <div class="kpi"><div class="label">Signed Today</div><div class="value">${signedToday}</div><div class="sub">of ${approved.length} total signed</div></div>
      <div class="kpi blue"><div class="label">Avg Wait Time</div><div class="value">${avgWait.toFixed(1)}h</div><div class="sub">Upload → Signature</div></div>
      <div class="kpi red"><div class="label">Oldest Pending</div><div class="value">${Math.floor(oldestPending)}h</div><div class="sub">${critical} over 48h</div></div>
    </div>
    ${storeEntries.length ? `
      <div class="chart-box">
        <h3>📊 Pending documents by store (top 10)</h3>
        <div class="bar-chart">
          ${storeEntries.map(([store, count]) => `<div class="bar" style="height:${(count/maxCount*140)}px" title="Store ${store}: ${count} pending"><span class="val">${count}</span><span class="lbl">${h(store)}</span></div>`).join('')}
        </div>
        <div style="height:24px;"></div>
      </div>` : ''}
    <div class="box">
      <h3>Recent activity</h3>
      <table><thead><tr><th>Doc</th><th>Store</th><th>Status</th><th>Age</th></tr></thead><tbody>
        ${all.slice().sort((a,b) => new Date(b.created_at) - new Date(a.created_at)).slice(0, 10).map(d => `<tr>
          <td>${new Date(d.created_at).toLocaleString()}</td>
          <td>${h(d.store_number || '-')}</td>
          <td><span class="badge ${d.status}">${d.status}</span></td>
          <td>${ageBadge(ageHours(d.created_at, d.signed_at))}</td>
        </tr>`).join('') || '<tr><td colspan="4" class="empty">No documents yet.</td></tr>'}
      </tbody></table>
    </div>`;
}

// Map of userId -> [{id, store_number, store_name}]
async function fetchStoreAssignments(userIds) {
  const map = {};
  if (!userIds.length) return map;
  const { data } = await sb.from('pgds_user_stores')
    .select('user_id, pgds_stores(id, store_number, store_name)')
    .in('user_id', userIds);
  (data || []).forEach(r => {
    if (!r.pgds_stores) return;
    (map[r.user_id] = map[r.user_id] || []).push(r.pgds_stores);
  });
  Object.values(map).forEach(list => list.sort((a, b) => String(a.store_number).localeCompare(String(b.store_number), undefined, { numeric: true })));
  return map;
}

function storeChips(list) {
  if (!list || !list.length) return '<span style="color:var(--muted);font-size:12px;">none assigned</span>';
  return list.map(s => `<span class="store-chip">${h(s.store_number)} - ${h(s.store_name)}</span>`).join('');
}

async function renderPendingUsers(el) {
  el.innerHTML = '<p class="empty">Loading...</p>';
  const { data, error } = await sb.from('pgds_profiles').select('*').eq('status', 'pending').order('created_at', { ascending: false });
  if (error) return el.innerHTML = `<div class="msg error">${error.message}</div>`;
  if (!data.length) return el.innerHTML = '<p class="empty">No pending approvals.</p>';
  const assigns = await fetchStoreAssignments(data.map(u => u.id));
  el.innerHTML = `<table><thead><tr><th>Email</th><th>Full Name</th><th>Assigned stores</th><th>Signed up</th><th>Actions</th></tr></thead><tbody>
    ${data.map(u => `<tr>
      <td>${h(u.email)}</td><td>${h(u.full_name||'-')}</td>
      <td>
        <div class="store-chips">${storeChips(assigns[u.id])}</div>
        <button class="secondary" style="margin-top:6px;padding:4px 10px;font-size:11px;" onclick="openStoreAssign('${u.id}')">✏️ Edit stores</button>
      </td>
      <td>${new Date(u.created_at).toLocaleString()}</td>
      <td class="actions">
        <button onclick="approveUser('${u.id}')">Approve</button>
        <button class="danger" onclick="rejectUser('${u.id}')">Reject</button>
      </td>
    </tr>`).join('')}</tbody></table>`;
}

// Dialog to assign stores to a user
async function openStoreAssign(userId) {
  closePopover();
  const [{ data: allStores }, assigns] = await Promise.all([
    sb.from('pgds_stores').select('*').eq('active', true).order('store_number', { ascending: true }),
    fetchStoreAssignments([userId])
  ]);
  const mine = new Set((assigns[userId] || []).map(s => s.id));
  const back = document.createElement('div'); back.className = 'popover-back'; back.onclick = closePopover;
  const pop = document.createElement('div'); pop.className = 'popover';
  pop.innerHTML = `
    <button class="popover-close" onclick="closePopover()">&times;</button>
    <h4>Assign stores</h4>
    <p style="font-size:12px;color:var(--muted);margin-bottom:10px;">Tick every store this person handles. They will choose which store each upload belongs to.</p>
    <div class="store-picker" id="assignPicker">
      ${(allStores || []).map(s => `
        <label class="store-opt">
          <input type="checkbox" value="${s.id}" data-num="${h(s.store_number)}" data-name="${h(s.store_name)}" ${mine.has(s.id) ? 'checked' : ''}>
          <span>${h(s.store_number)} - ${h(s.store_name)}</span>
        </label>`).join('')}
    </div>
    <div style="display:flex;gap:8px;margin-top:12px;">
      <button onclick="saveStoreAssign('${userId}')">💾 Save stores</button>
      <button class="secondary" onclick="closePopover()">Cancel</button>
    </div>`;
  document.body.appendChild(back); document.body.appendChild(pop);
}

async function saveStoreAssign(userId) {
  const picked = Array.from(document.querySelectorAll('#assignPicker input:checked'));
  if (!picked.length) return toast('Tick at least one store.', 'warn');
  try {
    await sb.from('pgds_user_stores').delete().eq('user_id', userId);
    const { error } = await sb.from('pgds_user_stores').insert(picked.map(p => ({ user_id: userId, store_id: p.value })));
    if (error) throw error;
    // Keep the profile's primary store in sync for display
    await sb.from('pgds_profiles').update({
      store_number: picked[0].dataset.num, store_name: picked[0].dataset.name
    }).eq('id', userId);
    closePopover();
    toast(`✓ ${picked.length} store(s) assigned`);
    renderAdminView();
  } catch (err) { toast(err.message || err, 'error'); }
}

async function renderAllUsers(el) {
  el.innerHTML = '<p class="empty">Loading...</p>';
  const { data, error } = await sb.from('pgds_profiles').select('*').order('created_at', { ascending: false });
  if (error) return el.innerHTML = `<div class="msg error">${error.message}</div>`;
  const nameById = {}; data.forEach(u => nameById[u.id] = u.full_name || u.email);
  const assigns = await fetchStoreAssignments(data.map(u => u.id));
  el.innerHTML = `<table><thead><tr><th>Email</th><th>Full Name</th><th>Assigned stores</th><th>Role</th><th>Status</th><th>Approved By</th><th>Actions</th></tr></thead><tbody>
    ${data.map(u => `<tr>
      <td>${h(u.email)}</td><td>${h(u.full_name||'-')}</td>
      <td>${u.role === 'admin' ? '<span style="color:var(--muted);font-size:12px;">all stores</span>' : `<div class="store-chips">${storeChips(assigns[u.id])}</div>`}</td>
      <td><span class="badge ${u.role==='admin'?'urgent':'normal'}">${u.role}</span></td>
      <td><span class="badge ${u.status}">${u.status}</span></td>
      <td>${u.approved_by&&nameById[u.approved_by]?h(nameById[u.approved_by]):'-'}${u.approved_at?`<div style="font-size:11px;color:var(--muted);">${new Date(u.approved_at).toLocaleDateString()}</div>`:''}</td>
      <td class="actions">
        ${u.role!=='admin'?`<button class="secondary" onclick="openStoreAssign('${u.id}')">✏️ Stores</button>`:''}
        ${u.status!=='approved'?`<button onclick="approveUser('${u.id}')">Approve</button>`:''}
        ${u.status!=='rejected'&&u.id!==currentUser.id?`<button class="danger" onclick="rejectUser('${u.id}')">Reject</button>`:''}
      </td>
    </tr>`).join('')}</tbody></table>`;
}

async function renderPendingQueue(el) {
  el.innerHTML = '<p class="empty">Loading signing queue...</p>';
  const { data, error } = await sb.from('pgds_documents').select('*').eq('status', 'pending').order('urgency', { ascending: false }).order('created_at', { ascending: false });
  if (error) return el.innerHTML = `<div class="msg error">${error.message}</div>`;
  window.__pgdsDocsCache = data;
  if (!data.length) return el.innerHTML = `<p class="empty">📭 No pending documents. All caught up!</p>`;
  const uids = [...new Set(data.map(d => d.uploader_id))];
  const profMap = {};
  if (uids.length) {
    const { data: ps } = await sb.from('pgds_profiles').select('id,full_name,email,store_number,store_name').in('id', uids);
    (ps||[]).forEach(p => profMap[p.id] = p);
  }
  el.innerHTML = `<table><thead><tr><th>Title</th><th>Uploader</th><th>Store</th><th>Urgency</th><th>Uploaded</th><th>Age</th><th>Remarks / Comments</th><th>Actions</th></tr></thead><tbody>
    ${data.map(d => { const up = profMap[d.uploader_id]||{}; const age = ageHours(d.created_at); return `<tr>
      <td><strong>${h(d.title)}</strong>${d.notes?`<div style="font-size:12px;color:var(--muted);margin-top:4px;">${h(d.notes)}</div>`:''}${d.signed_file_path?`<div style="font-size:11px;color:#059669;margin-top:4px;">✓ Partially signed - open to continue</div>`:''}</td>
      <td>${h(up.full_name||up.email||'-')}</td>
      <td>${h(d.store_number||up.store_number||'-')} - ${h(d.store_name||up.store_name||'')}</td>
      <td><span class="badge ${d.urgency}">${d.urgency}</span></td>
      <td>${new Date(d.created_at).toLocaleString()}</td>
      <td>${ageBadge(age)}</td>
      <td class="remarks-cell">${renderRemarksCell(d)}</td>
      <td class="actions">
        <button onclick="signDoc('${d.id}')">✍️ View & Sign</button>
        <button class="danger" onclick="rejectDoc('${d.id}')">Reject</button>
      </td>
    </tr>`; }).join('')}</tbody></table>`;
}

function renderRemarksCell(d) {
  if (d.admin_remarks) {
    return `<div class="remark-box">
      <div class="remark-text">${h(d.admin_remarks)}</div>
      <div class="remark-meta">${d.remarks_at ? new Date(d.remarks_at).toLocaleString() : ''}</div>
      <div class="remark-actions">
        <button class="secondary" onclick="openRemarksDialog('${d.id}')">✏️ Edit</button>
        <button class="danger" onclick="deleteRemarks('${d.id}')">🗑 Delete</button>
      </div>
    </div>`;
  }
  return `<button class="secondary" onclick="openRemarksDialog('${d.id}')">➕ Add remark</button>`;
}

function openRemarksDialog(docId) {
  closePopover();
  const row = (window.__pgdsDocsCache || []).find(x => x.id === docId);
  const existing = row ? (row.admin_remarks || '') : '';
  const back = document.createElement('div'); back.className = 'popover-back'; back.onclick = closePopover;
  const pop = document.createElement('div'); pop.className = 'popover';
  pop.innerHTML = `
    <button class="popover-close" onclick="closePopover()">&times;</button>
    <h4>${existing ? 'Edit remark' : 'Add remark for staff'}</h4>
    <p style="font-size:12px;color:var(--muted);margin-bottom:8px;">The uploader will see this on their dashboard - e.g. missing requirements.</p>
    <textarea id="remarkText" style="min-height:110px;font-size:14px;" placeholder="e.g. Missing the area manager signature on page 2. Please re-upload.">${h(existing)}</textarea>
    <div style="display:flex;gap:8px;margin-top:10px;">
      <button onclick="saveRemarks('${docId}')">💾 Save remark</button>
      <button class="secondary" onclick="closePopover()">Cancel</button>
    </div>`;
  document.body.appendChild(back); document.body.appendChild(pop);
  setTimeout(() => { const t = document.getElementById('remarkText'); if (t) { t.focus(); t.setSelectionRange(t.value.length, t.value.length); } }, 50);
}

async function saveRemarks(docId) {
  const text = document.getElementById('remarkText').value.trim();
  if (!text) return toast('Type a remark first.', 'warn');
  const { error } = await sb.from('pgds_documents').update({
    admin_remarks: text, remarks_at: new Date().toISOString(), remarks_by: currentUser.id
  }).eq('id', docId);
  if (error) return toast(error.message, 'error');
  closePopover();
  toast('✓ Remark saved');
  renderAdminView();
}

async function deleteRemarks(docId) {
  if (!confirm('Delete this remark?')) return;
  const { error } = await sb.from('pgds_documents').update({
    admin_remarks: null, remarks_at: null, remarks_by: null
  }).eq('id', docId);
  if (error) return toast(error.message, 'error');
  toast('Remark deleted');
  renderAdminView();
}

async function renderRejected(el) {
  el.innerHTML = '<p class="empty">Loading rejected...</p>';
  const { data, error } = await sb.from('pgds_documents').select('*').eq('status', 'rejected').order('created_at', { ascending: false });
  if (error) return el.innerHTML = `<div class="msg error">${error.message}</div>`;
  if (!data.length) return el.innerHTML = `<p class="empty">No rejected documents.</p>`;
  const uids = [...new Set(data.map(d => d.uploader_id))];
  const profMap = {};
  if (uids.length) {
    const { data: ps } = await sb.from('pgds_profiles').select('id,full_name,email,store_number,store_name').in('id', uids);
    (ps||[]).forEach(p => profMap[p.id] = p);
  }
  el.innerHTML = `<table><thead><tr><th>Title</th><th>Uploader</th><th>Store</th><th>Reason</th><th>Rejected</th><th>Actions</th></tr></thead><tbody>
    ${data.map(d => { const up = profMap[d.uploader_id]||{}; return `<tr>
      <td><strong>${h(d.title)}</strong></td>
      <td>${h(up.full_name||up.email||'-')}</td>
      <td>${h(d.store_number||up.store_number||'-')} - ${h(d.store_name||up.store_name||'')}</td>
      <td style="color:#991b1b;">${h(d.reject_reason||'-')}</td>
      <td>${d.signed_at?new Date(d.signed_at).toLocaleString():'-'}</td>
      <td class="actions">
        <button class="secondary" onclick="viewAdminDoc('pgds-pending','${d.original_file_path}')">👁️ View</button>
        <button class="danger" onclick="deleteApprovedDoc('${d.id}','${d.original_file_path}','${d.signed_file_path||''}')">🗑 Delete</button>
      </td>
    </tr>`; }).join('')}</tbody></table>`;
}

async function renderApprovedFolders(el) {
  el.innerHTML = '<p class="empty">Loading folders...</p>';
  const { data } = await sb.from('pgds_documents').select('*').eq('status', 'approved').order('signed_at', { ascending: false });
  if (!data || !data.length) return el.innerHTML = '<p class="empty">No approved documents yet.</p>';
  const uids = [...new Set(data.map(d => d.uploader_id))];
  const profMap = {};
  if (uids.length) {
    const { data: ps } = await sb.from('pgds_profiles').select('id,full_name,email,store_number,store_name').in('id', uids);
    (ps||[]).forEach(p => profMap[p.id] = p);
  }
  // Group by store
  const byStore = {};
  data.forEach(d => {
    const up = profMap[d.uploader_id] || {};
    const key = `${d.store_number || up.store_number || '?'} - ${d.store_name || up.store_name || 'Unassigned'}`;
    (byStore[key] = byStore[key] || []).push({ d, up });
  });
  el.innerHTML = `<h3 style="margin-bottom:12px;color:var(--green-900);">📁 Approved documents by store</h3>` +
    Object.entries(byStore).sort().map(([store, items], idx) => `
      <div class="folder-group ${idx===0?'open':''}">
        <div class="folder-head" onclick="this.parentElement.classList.toggle('open')">
          <span class="name">📂 ${h(store)}</span>
          <span class="count">${items.length}</span>
        </div>
        <div class="folder-body">
          <table><thead><tr><th>Title</th><th>Uploaded by</th><th>Signed</th><th>Actions</th></tr></thead><tbody>
            ${items.map(({d, up}) => `<tr>
              <td><strong>${h(d.title)}</strong></td>
              <td>${h(up.full_name || up.email || '-')}</td>
              <td>${d.signed_at ? new Date(d.signed_at).toLocaleString() : '-'}</td>
              <td class="actions">
                <button onclick="downloadDirect('pgds-signed','${d.signed_file_path}','${h(safeFileName(d.title)).replace(/'/g,"\\'")}')">⬇️ Download</button>
                <button class="secondary" onclick="viewAdminDoc('pgds-signed','${d.signed_file_path}')">👁️ View</button>
                <button onclick="signDoc('${d.id}')">✏️ Edit</button>
                <button class="danger" onclick="deleteApprovedDoc('${d.id}','${d.original_file_path}','${d.signed_file_path}')">🗑 Delete</button>
              </td>
            </tr>`).join('')}
          </tbody></table>
        </div>
      </div>`).join('');
}

async function deleteApprovedDoc(id, originalPath, signedPath) {
  if (!confirm('Delete this document PERMANENTLY?\n\nIt will disappear for the uploader too. This cannot be undone.')) return;
  try {
    // Delete the DB row first and confirm a row actually came back.
    // Under RLS a blocked delete returns success with zero rows, so we must check.
    const { data: deleted, error } = await sb.from('pgds_documents').delete().eq('id', id).select('id');
    if (error) throw error;
    if (!deleted || !deleted.length) {
      throw new Error('Nothing was deleted - the database is blocking it. Run the admin delete policy SQL in Supabase.');
    }
    // Row is gone for everyone; now clean up the files (best effort)
    if (signedPath) { try { await sb.storage.from('pgds-signed').remove([signedPath]); } catch (e) {} }
    if (originalPath) { try { await sb.storage.from('pgds-pending').remove([originalPath]); } catch (e) {} }
    toast('Document deleted - removed for the uploader too');
    renderAdminView(); updateAdminCounts();
  } catch (err) { toast('Delete failed: ' + (err.message || err), 'error'); }
}

async function viewAdminDoc(bucket, path) {
  await openInAppViewer(bucket, path);
}

async function rejectDoc(id) {
  const reason = prompt('Reason for rejection:') || '';
  if (!reason) return;
  const { error } = await sb.from('pgds_documents').update({ status: 'rejected', reject_reason: reason, signed_by: currentUser.id, signed_at: new Date().toISOString() }).eq('id', id);
  if (error) return showMsg('msg', error.message, 'error');
  showMsg('msg', 'Document rejected.', 'success');
  renderAdminView(); updateAdminCounts();
}

async function approveUser(id) {
  // Make sure they have at least one store assigned before approving
  const assigns = await fetchStoreAssignments([id]);
  if (!assigns[id] || !assigns[id].length) {
    toast('Assign at least one store first.', 'warn');
    return openStoreAssign(id);
  }
  const { error } = await sb.from('pgds_profiles').update({
    status: 'approved', approved_at: new Date().toISOString(), approved_by: currentUser.id
  }).eq('id', id);
  if (error) return toast(error.message, 'error');
  toast('✓ User approved');
  renderAdminView(); updateAdminCounts();
}

async function rejectUser(id) {
  const reason = prompt('Reason for rejection (optional):') || '';
  const { error } = await sb.from('pgds_profiles').update({ status: 'rejected', reject_reason: reason }).eq('id', id);
  if (error) return showMsg('msg', error.message, 'error');
  showMsg('msg', 'User rejected.', 'success');
  renderAdminView(); updateAdminCounts();
}

async function renderStores(el) {
  el.innerHTML = `
    <div class="box">
      <h3>Add a new store</h3>
      <form id="addStoreForm" style="display:grid;grid-template-columns:120px 1fr auto;gap:10px;align-items:end;">
        <div class="form-group" style="margin:0"><label>Store #</label><input type="text" id="newStoreNum" required></div>
        <div class="form-group" style="margin:0"><label>Store name</label><input type="text" id="newStoreName" required></div>
        <button type="submit">Add store</button>
      </form>
    </div>
    <div id="storesBody"><p class="empty">Loading...</p></div>`;
  document.getElementById('addStoreForm').onsubmit = async (e) => {
    e.preventDefault();
    const num = document.getElementById('newStoreNum').value.trim();
    const name = document.getElementById('newStoreName').value.trim();
    const { error } = await sb.from('pgds_stores').insert({ store_number: num, store_name: name, created_by: currentUser.id });
    if (error) return showMsg('msg', error.message, 'error');
    showMsg('msg', 'Store added.', 'success'); renderStores(el);
  };
  const body = document.getElementById('storesBody');
  const { data, error } = await sb.from('pgds_stores').select('*').order('store_number', { ascending: true });
  if (error) return body.innerHTML = `<div class="msg error">${error.message}</div>`;
  if (!data.length) return body.innerHTML = '<p class="empty">No stores yet.</p>';
  body.innerHTML = `<table><thead><tr><th>Store #</th><th>Store Name / Branch</th><th>Status</th><th>Added</th><th>Actions</th></tr></thead><tbody>
    ${data.map(s => `<tr>
      <td><input type="text" style="width:80px" id="stn-${s.id}" value="${h(s.store_number)}"></td>
      <td><input type="text" style="width:100%" id="stnm-${s.id}" value="${h(s.store_name)}"></td>
      <td><span class="badge ${s.active?'approved':'rejected'}">${s.active?'active':'inactive'}</span></td>
      <td>${new Date(s.created_at).toLocaleDateString()}</td>
      <td class="actions">
        <button onclick="saveStore('${s.id}')">Save</button>
        <button class="secondary" onclick="toggleStore('${s.id}',${!s.active})">${s.active?'Deactivate':'Activate'}</button>
        <button class="danger" onclick="deleteStore('${s.id}')">Delete</button>
      </td>
    </tr>`).join('')}</tbody></table>`;
}

async function saveStore(id) {
  const num = document.getElementById(`stn-${id}`).value.trim();
  const name = document.getElementById(`stnm-${id}`).value.trim();
  const { error } = await sb.from('pgds_stores').update({ store_number: num, store_name: name }).eq('id', id);
  if (error) return showMsg('msg', error.message, 'error');
  showMsg('msg', 'Store updated.', 'success');
}
async function toggleStore(id, active) {
  const { error } = await sb.from('pgds_stores').update({ active }).eq('id', id);
  if (error) return showMsg('msg', error.message, 'error');
  renderAdminView();
}
async function deleteStore(id) {
  if (!confirm('Delete this store?')) return;
  const { error } = await sb.from('pgds_stores').delete().eq('id', id);
  if (error) return showMsg('msg', error.message, 'error');
  renderAdminView();
}

async function renderSignaturesMgr(el) {
  el.innerHTML = '<p class="empty">Loading...</p>';
  const { data } = await sb.from('pgds_signatures').select('*').eq('owner_id', currentUser.id).order('created_at', { ascending: false });
  el.innerHTML = `
    <div class="box">
      <h3>Draw a new saved signature</h3>
      <p style="font-size:13px;color:var(--muted);margin-bottom:10px;">Draw once, reuse forever from the Sign panel.</p>
      <canvas id="sigMgrCanvas" class="draw-canvas" width="400" height="140"></canvas>
      <div style="display:flex;gap:8px;margin-top:10px;align-items:end;">
        <div style="flex:1;"><label>Name</label><input type="text" id="sigMgrName" placeholder="e.g. Main signature"></div>
        <button class="secondary" onclick="clearMgrCanvas()">Clear</button>
        <button onclick="saveMgrSig()">Save signature</button>
      </div>
    </div>
    <h3 style="color:var(--green-900);margin-bottom:10px;">Saved signatures</h3>
    ${(data||[]).length ? `<div class="saved-sigs" style="grid-template-columns:repeat(auto-fill,minmax(160px,1fr));">
      ${(data||[]).map(s => `<div class="saved-sig">
        <img src="${s.image_data}">
        <div style="font-size:12px;margin-top:6px;">${h(s.name)}</div>
        <span class="del" onclick="deleteSavedSigMgr('${s.id}')">&times;</span>
      </div>`).join('')}
    </div>` : '<p class="empty">No saved signatures yet.</p>'}`;
  initMgrCanvas();
}

let mgr = { ctx: null, canvas: null, drawing: false, drawn: false, lastMid: null, currentStroke: [], dpr: 1, color: '#1e40af', width: 3 };
function initMgrCanvas() {
  const canvas = document.getElementById('sigMgrCanvas');
  const rect = canvas.getBoundingClientRect();
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.floor(rect.width * dpr);
  canvas.height = Math.floor(rect.height * dpr);
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.lineCap = 'round'; ctx.lineJoin = 'round';
  mgr = { ctx, canvas, drawing: false, drawn: false, lastMid: null, currentStroke: [], dpr, color: '#1e40af', width: 3 };
  const getPos = (e) => { const r = canvas.getBoundingClientRect(); const t = e.touches ? e.touches[0] : e; return { x: t.clientX - r.left, y: t.clientY - r.top }; };
  canvas.onmousedown = canvas.ontouchstart = (e) => { mgr.drawing = true; mgr.drawn = true; const p = getPos(e); mgr.currentStroke = [p]; mgr.lastMid = p; e.preventDefault(); };
  canvas.onmousemove = canvas.ontouchmove = (e) => {
    if (!mgr.drawing) return;
    const p = getPos(e);
    mgr.currentStroke.push(p);
    const prev = mgr.currentStroke[mgr.currentStroke.length - 2];
    if (prev) {
      const mid = { x: (prev.x + p.x) / 2, y: (prev.y + p.y) / 2 };
      ctx.strokeStyle = mgr.color; ctx.lineWidth = mgr.width;
      ctx.beginPath(); ctx.moveTo(mgr.lastMid.x, mgr.lastMid.y);
      ctx.quadraticCurveTo(prev.x, prev.y, mid.x, mid.y); ctx.stroke();
      mgr.lastMid = mid;
    }
    e.preventDefault();
  };
  canvas.onmouseup = canvas.ontouchend = canvas.onmouseleave = () => {
    if (!mgr.drawing) return;
    mgr.drawing = false;
    if (mgr.currentStroke.length >= 2) {
      const last = mgr.currentStroke[mgr.currentStroke.length - 1];
      ctx.beginPath(); ctx.moveTo(mgr.lastMid.x, mgr.lastMid.y); ctx.lineTo(last.x, last.y); ctx.stroke();
    }
    mgr.currentStroke = [];
  };
}
function clearMgrCanvas() { if (mgr.ctx) { mgr.ctx.clearRect(0,0, mgr.canvas.width / mgr.dpr, mgr.canvas.height / mgr.dpr); mgr.drawn = false; } }
async function saveMgrSig() {
  if (!mgr.drawn) return alert('Please draw a signature first.');
  const name = document.getElementById('sigMgrName').value.trim() || 'Signature';
  // Trim
  const c = mgr.canvas;
  const w = c.width, hh = c.height;
  const px = c.getContext('2d').getImageData(0, 0, w, hh).data;
  let top = hh, bottom = 0, left = w, right = 0;
  for (let y = 0; y < hh; y++) for (let x = 0; x < w; x++) {
    if (px[(y * w + x) * 4 + 3] > 0) {
      if (y < top) top = y; if (y > bottom) bottom = y;
      if (x < left) left = x; if (x > right) right = x;
    }
  }
  const pad2 = 10;
  const nw = right - left + pad2 * 2, nh = bottom - top + pad2 * 2;
  const out = document.createElement('canvas'); out.width = nw; out.height = nh;
  out.getContext('2d').drawImage(c, left - pad2, top - pad2, nw, nh, 0, 0, nw, nh);
  const dataUrl = out.toDataURL('image/png');
  const { error } = await sb.from('pgds_signatures').insert({ owner_id: currentUser.id, name, image_data: dataUrl });
  if (error) return showMsg('msg', error.message, 'error');
  showMsg('msg', 'Signature saved.', 'success');
  renderSignaturesMgr(document.getElementById('content'));
}
async function deleteSavedSigMgr(id) {
  if (!confirm('Delete this signature?')) return;
  await sb.from('pgds_signatures').delete().eq('id', id);
  renderSignaturesMgr(document.getElementById('content'));
}

// ---------- SIGNING ENGINE ----------
let sigState = { doc: null, pages: [], overlays: [], mySignatures: [], pdfBytes: null, pageInfo: [] };

async function signDoc(id) {
  const { data: docRow, error } = await sb.from('pgds_documents').select('*').eq('id', id).single();
  if (error) return alert(error.message);
  sigState = { doc: docRow, pages: [], pageWraps: [], overlays: [], mySignatures: [], pdfBytes: null, pageInfo: [], pageZoom: [] };
  const { data: sigs } = await sb.from('pgds_signatures').select('*').eq('owner_id', currentUser.id).order('created_at', { ascending: false });
  sigState.mySignatures = sigs || [];
  await openSigningModal();
}

async function openSigningModal() {
  const doc = sigState.doc;
  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.id = 'sigModal';
  modal.innerHTML = `
    <div class="modal-topbar">
      <span class="title">✍️ ${h(doc.title)}</span>
      <div style="display:flex;gap:6px;align-items:center;">
        <button onclick="zoomFit()" title="Reset all pages to 100%" style="padding:6px 10px;">↔️ Fit All</button>
        <button onclick="showOverlaysList()">Items: <span id="overlayCount">0</span></button>
        <button onclick="closeSigningModal()">Close</button>
      </div>
    </div>
    <div class="modal-doc" id="modalDoc"><p style="color:#fff">Loading document...</p></div>
    <div class="modal-bottombar">
      <button class="modal-tool-btn" onclick="openSignaturePad()"><span class="ic">✍️</span>Signature</button>
      <button class="modal-tool-btn" onclick="showSavedPopover()"><span class="ic">📎</span>Saved</button>
      <button class="modal-tool-btn" onclick="uploadSignatureImage()"><span class="ic">📤</span>Upload</button>
      <button class="modal-tool-btn" onclick="showTextPopover()"><span class="ic">🔤</span>Add Text</button>
      <button class="modal-tool-btn" onclick="showStampPopover()"><span class="ic">✅</span>Stamp</button>
      <button class="modal-tool-btn" id="saveBtn" onclick="saveSignatures()"><span class="ic">💾</span>Save</button>
      <button class="modal-tool-btn primary" id="approveBtn" onclick="approveDoc()"><span class="ic">✅</span>Approve</button>
    </div>
    <input type="file" id="sigUploadInput" accept="image/png,image/jpeg" style="display:none;">`;
  document.body.appendChild(modal);
  await loadDocForSigning();
}

function closePopover() {
  document.querySelectorAll('.popover, .popover-back').forEach(e => e.remove());
}

function showSavedPopover() {
  closePopover();
  const back = document.createElement('div'); back.className = 'popover-back'; back.onclick = closePopover;
  const pop = document.createElement('div'); pop.className = 'popover';
  pop.innerHTML = `
    <button class="popover-close" onclick="closePopover()">&times;</button>
    <h4>Saved signatures</h4>
    ${sigState.mySignatures.length ? `<div class="saved-sigs">
      ${sigState.mySignatures.map(s => `<div class="saved-sig" onclick="addSavedToDoc('${s.id}');closePopover();">
        <span class="del" onclick="event.stopPropagation();deleteSavedSig('${s.id}')">&times;</span>
        <img src="${s.image_data}"><div style="font-size:11px;margin-top:4px;">${h(s.name)}</div>
      </div>`).join('')}
    </div>` : '<p style="font-size:13px;color:var(--muted);">No saved signatures yet. Tap Signature and save one from the pad.</p>'}`;
  document.body.appendChild(back); document.body.appendChild(pop);
}

function showTextPopover() {
  closePopover();
  const back = document.createElement('div'); back.className = 'popover-back'; back.onclick = closePopover;
  const pop = document.createElement('div'); pop.className = 'popover';
  pop.innerHTML = `
    <button class="popover-close" onclick="closePopover()">&times;</button>
    <h4>Type a name</h4>
    <p style="font-size:12px;color:var(--muted);margin-bottom:8px;">Useful when staff forgot to write the signer's name.</p>
    <input type="text" id="typedName" placeholder="e.g. C. M. Nemo" style="font-size:15px;margin-bottom:10px;">
    <button onclick="addTypedToDoc();closePopover();">Add to doc</button>`;
  document.body.appendChild(back); document.body.appendChild(pop);
  setTimeout(() => document.getElementById('typedName').focus(), 50);
}

function showStampPopover() {
  closePopover();
  const back = document.createElement('div'); back.className = 'popover-back'; back.onclick = closePopover;
  const pop = document.createElement('div'); pop.className = 'popover';
  pop.innerHTML = `
    <button class="popover-close" onclick="closePopover()">&times;</button>
    <h4>Add approval stamp</h4>
    <select id="stampText" style="margin-bottom:10px;">
      <option>APPROVED</option><option>REVIEWED</option><option>PAID</option><option>RECEIVED</option>
    </select>
    <button onclick="addStampToDoc();closePopover();">Add to doc</button>`;
  document.body.appendChild(back); document.body.appendChild(pop);
}

function uploadSignatureImage() {
  const inp = document.getElementById('sigUploadInput');
  inp.onchange = async () => {
    const f = inp.files[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      addImageOverlay(reader.result, 200, 70);
      inp.value = '';
      toast('Signature added - drag to position, then tap Save');
    };
    reader.readAsDataURL(f);
  };
  inp.click();
}

function showOverlaysList() {
  closePopover();
  const back = document.createElement('div'); back.className = 'popover-back'; back.onclick = closePopover;
  const pop = document.createElement('div'); pop.className = 'popover';
  pop.innerHTML = `
    <button class="popover-close" onclick="closePopover()">&times;</button>
    <h4>Placed items on document</h4>
    ${sigState.overlays.length ? sigState.overlays.map((o, i) => `<div style="padding:8px 0;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;">
      <span>${i+1}. ${o.type === 'image' ? '✍️ Signature' : (o.isStamp ? '✅ Stamp: ' + o.text : '🔤 ' + o.text)}</span>
      <span style="color:var(--red);cursor:pointer;font-weight:700;" onclick="removeOverlay('${o.id}');showOverlaysList();">Remove</span>
    </div>`).join('') : '<p style="font-size:13px;color:var(--muted);">Nothing placed yet.</p>'}`;
  document.body.appendChild(back); document.body.appendChild(pop);
}

// ---------- SIGNATURE PAD (fullscreen, landscape, smooth ink) ----------
let pad = { ctx: null, canvas: null, strokes: [], currentStroke: [], drawing: false, color: '#1e40af', width: 3, lastMid: null, dpr: 1 };

function openSignaturePad() {
  const wrap = document.createElement('div');
  wrap.id = 'sigPad';
  wrap.className = 'sigpad';
  wrap.innerHTML = `
    <div class="sigpad-header">
      <span class="title"><button class="back" onclick="closeSignaturePad()">←</button>Signature Pad</span>
      <div class="tools">
        <button class="trash" onclick="clearPad()" title="Clear">🗑 Clear</button>
        <button class="save" onclick="savePadSignature()">SAVE SIGN</button>
      </div>
    </div>
    <div class="sigpad-canvas-wrap"><canvas id="padCanvas" class="sigpad-canvas"></canvas></div>
    <div class="sigpad-controls">
      <div class="color-swatches" id="colorSwatches">
        <div class="color-swatch active" style="background:#1e40af" onclick="setPadColor('#1e40af',this)"></div>
        <div class="color-swatch" style="background:#0f172a" onclick="setPadColor('#0f172a',this)"></div>
        <div class="color-swatch" style="background:#dc2626" onclick="setPadColor('#dc2626',this)"></div>
        <div class="color-swatch" style="background:#064e3b" onclick="setPadColor('#064e3b',this)"></div>
      </div>
      <div class="pen-thickness">
        <button class="pen-btn" onclick="setPadWidth(2,this)"><span class="dot" style="width:6px;height:6px;"></span>Thin</button>
        <button class="pen-btn active" onclick="setPadWidth(3,this)"><span class="dot" style="width:9px;height:9px;"></span>Medium</button>
        <button class="pen-btn" onclick="setPadWidth(5,this)"><span class="dot" style="width:13px;height:13px;"></span>Thick</button>
      </div>
      <label style="display:flex;gap:8px;align-items:center;font-size:13px;">
        <input type="checkbox" id="padSaveReuse" style="width:auto;"> Save for reuse
      </label>
      <input type="text" id="padSaveName" placeholder="Signature name (if saving)" style="width:180px;font-size:13px;">
    </div>`;
  document.body.appendChild(wrap);
  setTimeout(initPadCanvas, 50);
}

function initPadCanvas() {
  const canvas = document.getElementById('padCanvas');
  if (!canvas) return;
  const rect = canvas.getBoundingClientRect();
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.floor(rect.width * dpr);
  canvas.height = Math.floor(rect.height * dpr);
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.lineCap = 'round'; ctx.lineJoin = 'round';
  pad.ctx = ctx; pad.canvas = canvas; pad.strokes = []; pad.currentStroke = []; pad.dpr = dpr;
  pad.lastMid = null;

  const getPos = (e) => {
    const r = canvas.getBoundingClientRect();
    const t = e.touches ? e.touches[0] : e;
    return { x: t.clientX - r.left, y: t.clientY - r.top };
  };
  const start = (e) => {
    pad.drawing = true;
    const p = getPos(e);
    pad.currentStroke = [p];
    pad.lastMid = p;
    e.preventDefault();
  };
  const move = (e) => {
    if (!pad.drawing) return;
    const p = getPos(e);
    pad.currentStroke.push(p);
    // Draw quadratic curve using midpoints for smooth line
    const prev = pad.currentStroke[pad.currentStroke.length - 2];
    if (prev) {
      const mid = { x: (prev.x + p.x) / 2, y: (prev.y + p.y) / 2 };
      ctx.strokeStyle = pad.color;
      ctx.lineWidth = pad.width;
      ctx.beginPath();
      ctx.moveTo(pad.lastMid.x, pad.lastMid.y);
      ctx.quadraticCurveTo(prev.x, prev.y, mid.x, mid.y);
      ctx.stroke();
      pad.lastMid = mid;
    }
    e.preventDefault();
  };
  const end = (e) => {
    if (!pad.drawing) return;
    pad.drawing = false;
    // Finalize by drawing to the last point
    if (pad.currentStroke.length >= 2) {
      const last = pad.currentStroke[pad.currentStroke.length - 1];
      ctx.beginPath();
      ctx.moveTo(pad.lastMid.x, pad.lastMid.y);
      ctx.lineTo(last.x, last.y);
      ctx.stroke();
    }
    pad.strokes.push({ points: pad.currentStroke, color: pad.color, width: pad.width });
    pad.currentStroke = [];
    if (e) e.preventDefault();
  };
  canvas.onmousedown = start; canvas.onmousemove = move; canvas.onmouseup = end; canvas.onmouseleave = end;
  canvas.ontouchstart = start; canvas.ontouchmove = move; canvas.ontouchend = end;
}

function setPadColor(c, el) {
  pad.color = c;
  document.querySelectorAll('#colorSwatches .color-swatch').forEach(x => x.classList.remove('active'));
  el.classList.add('active');
}
function setPadWidth(w, el) {
  pad.width = w;
  document.querySelectorAll('.pen-btn').forEach(x => x.classList.remove('active'));
  el.classList.add('active');
}

function clearPad() {
  if (!pad.ctx) return;
  const c = pad.canvas;
  pad.ctx.clearRect(0, 0, c.width / pad.dpr, c.height / pad.dpr);
  pad.strokes = [];
}

function closeSignaturePad() {
  const p = document.getElementById('sigPad'); if (p) p.remove();
}

async function savePadSignature() {
  if (!pad.strokes.length) return toast('Draw your signature first.', 'warn');
  const trimmed = trimPadCanvas();
  const dataUrl = trimmed.toDataURL('image/png');
  // Add to current document
  addImageOverlay(dataUrl, 200, 70);
  // Save if requested
  if (document.getElementById('padSaveReuse').checked) {
    const name = document.getElementById('padSaveName').value.trim() || 'Signature ' + (sigState.mySignatures.length + 1);
    const { data } = await sb.from('pgds_signatures').insert({ owner_id: currentUser.id, name, image_data: dataUrl }).select().single();
    if (data) sigState.mySignatures.unshift(data);
  }
  closeSignaturePad();
}

function trimPadCanvas() {
  const c = pad.canvas;
  const w = c.width, hh = c.height;
  const ctx = c.getContext('2d');
  const px = ctx.getImageData(0, 0, w, hh).data;
  let top = hh, bottom = 0, left = w, right = 0;
  for (let y = 0; y < hh; y++) for (let x = 0; x < w; x++) {
    if (px[(y * w + x) * 4 + 3] > 0) {
      if (y < top) top = y; if (y > bottom) bottom = y;
      if (x < left) left = x; if (x > right) right = x;
    }
  }
  if (right < left) { const empty = document.createElement('canvas'); empty.width = 200; empty.height = 60; return empty; }
  const pad2 = 10;
  const nw = right - left + pad2 * 2, nh = bottom - top + pad2 * 2;
  const out = document.createElement('canvas'); out.width = nw; out.height = nh;
  out.getContext('2d').drawImage(c, left - pad2, top - pad2, nw, nh, 0, 0, nw, nh);
  return out;
}

function trimCanvas(c) {
  const ctx = c.getContext('2d');
  const w = c.width, hh = c.height;
  const px = ctx.getImageData(0, 0, w, hh).data;
  let top = hh, bottom = 0, left = w, right = 0;
  for (let y = 0; y < hh; y++) for (let x = 0; x < w; x++) {
    if (px[(y * w + x) * 4 + 3] > 0) {
      if (y < top) top = y; if (y > bottom) bottom = y;
      if (x < left) left = x; if (x > right) right = x;
    }
  }
  if (right < left) return c;
  const nw = right - left + 8, nh = bottom - top + 8;
  const out = document.createElement('canvas'); out.width = nw; out.height = nh;
  out.getContext('2d').drawImage(c, left - 4, top - 4, nw, nh, 0, 0, nw, nh);
  return out;
}

function addSavedToDoc(id) {
  const s = sigState.mySignatures.find(x => x.id === id);
  if (s) addImageOverlay(s.image_data, 180, 60);
}
function addTypedToDoc() {
  const t = document.getElementById('typedName').value.trim();
  if (!t) return toast('Type a name first.', 'warn');
  addTextOverlay(t, false);
}
function addStampToDoc() {
  const t = document.getElementById('stampText').value;
  addTextOverlay(t, true);
}
async function deleteSavedSig(id) {
  if (!confirm('Delete this saved signature?')) return;
  await sb.from('pgds_signatures').delete().eq('id', id);
  sigState.mySignatures = sigState.mySignatures.filter(s => s.id !== id);
  renderToolPanel('saved');
}

let overlayIdCounter = 0;
function addImageOverlay(dataUrl, w, hh) {
  if (!sigState.pages.length) return toast('Document not ready.', 'warn');
  const target = visiblePageAndPos(w, hh);
  const id = 'ov-' + (++overlayIdCounter);
  const div = document.createElement('div');
  div.className = 'sig-overlay'; div.id = id;
  div.style.left = target.left + 'px'; div.style.top = target.top + 'px';
  div.style.width = w + 'px'; div.style.height = hh + 'px';
  div.innerHTML = `<img src="${dataUrl}"><span class="remove-handle" onclick="removeOverlay('${id}')">&times;</span><span class="resize-handle"></span>`;
  target.pageWrap.appendChild(div);
  makeDraggable(div);
  sigState.overlays.push({ id, type: 'image', dataUrl, el: div, pageIdx: target.pageIdx });
  updateOverlayList();
  setTimeout(() => div.scrollIntoView({ behavior: 'smooth', block: 'center' }), 50);
}
function addTextOverlay(text, isStamp) {
  if (!sigState.pages.length) return toast('Document not ready.', 'warn');
  const w = isStamp ? 180 : Math.max(120, text.length * 15);
  const hh = isStamp ? 60 : 50;
  const target = visiblePageAndPos(w, hh);
  const id = 'ov-' + (++overlayIdCounter);
  const div = document.createElement('div');
  div.className = 'sig-overlay'; div.id = id;
  div.style.left = target.left + 'px'; div.style.top = target.top + 'px';
  div.style.width = w + 'px'; div.style.height = hh + 'px';
  div.innerHTML = `<div class="sig-text ${isStamp?'stamp':''}">${h(text)}</div><span class="remove-handle" onclick="removeOverlay('${id}')">&times;</span><span class="resize-handle"></span>`;
  target.pageWrap.appendChild(div);
  makeDraggable(div);
  sigState.overlays.push({ id, type: 'text', text, isStamp, el: div, pageIdx: target.pageIdx });
  updateOverlayList();
  setTimeout(() => div.scrollIntoView({ behavior: 'smooth', block: 'center' }), 50);
}
function removeOverlay(id) {
  const ov = sigState.overlays.find(o => o.id === id);
  if (ov) { ov.el.remove(); sigState.overlays = sigState.overlays.filter(o => o.id !== id); updateOverlayList(); }
}
function updateOverlayList() {
  const cnt = document.getElementById('overlayCount');
  if (cnt) cnt.textContent = sigState.overlays.length;
}

function makeDraggable(el) {
  let dragging = false, resizing = false, sx = 0, sy = 0, sl = 0, st = 0, sw = 0, sh = 0;
  const start = (e) => {
    const t = e.touches ? e.touches[0] : e;
    if (e.target.classList.contains('resize-handle')) { resizing = true; sw = el.offsetWidth; sh = el.offsetHeight; }
    else if (e.target.classList.contains('remove-handle')) return;
    else { dragging = true; }
    sx = t.clientX; sy = t.clientY; sl = el.offsetLeft; st = el.offsetTop;
    e.preventDefault();
  };
  const move = (e) => {
    if (!dragging && !resizing) return;
    const t = e.touches ? e.touches[0] : e;
    // Determine which page this overlay belongs to and use its zoom
    const ov = sigState.overlays.find(o => o.el === el);
    const z = ov ? pageZoom(ov.pageIdx) : 1;
    const dx = (t.clientX - sx) / z, dy = (t.clientY - sy) / z;
    if (dragging) { el.style.left = (sl + dx) + 'px'; el.style.top = (st + dy) + 'px'; }
    else if (resizing) {
      const nw = Math.max(40, sw + dx), nh = Math.max(20, sh + dy);
      el.style.width = nw + 'px'; el.style.height = nh + 'px';
      const t2 = el.querySelector('.sig-text'); if (t2) t2.style.fontSize = Math.max(12, nh * 0.55) + 'px';
    }
    e.preventDefault();
  };
  const end = () => { dragging = false; resizing = false; };
  el.onmousedown = start; el.ontouchstart = start;
  document.addEventListener('mousemove', move); document.addEventListener('touchmove', move, { passive: false });
  document.addEventListener('mouseup', end); document.addEventListener('touchend', end);
}

async function loadDocForSigning() {
  const doc = sigState.doc;
  const modalDoc = document.getElementById('modalDoc');
  // If there's already a signed version (saved partial or fully approved), load it
  const useSigned = !!doc.signed_file_path;
  const bucket = useSigned ? 'pgds-signed' : 'pgds-pending';
  const path = useSigned ? doc.signed_file_path : doc.original_file_path;
  const { data: signed, error: se } = await sb.storage.from(bucket).createSignedUrl(path, 300);
  if (se) return modalDoc.innerHTML = `<div class="msg error">${se.message}</div>`;
  const resp = await fetch(signed.signedUrl);
  const buf = await resp.arrayBuffer();
  const ext = path.split('.').pop().toLowerCase();
  let pdfBytes;
  if (ext === 'pdf') pdfBytes = new Uint8Array(buf);
  else {
    const pdfDoc = await PDFLib.PDFDocument.create();
    const imgBytes = new Uint8Array(buf);
    const img = ['jpg','jpeg'].includes(ext) ? await pdfDoc.embedJpg(imgBytes) : await pdfDoc.embedPng(imgBytes);
    const page = pdfDoc.addPage([img.width, img.height]);
    page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
    pdfBytes = await pdfDoc.save();
  }
  sigState.pdfBytes = pdfBytes;
  await renderPdfPages(pdfBytes);
  // Initialize per-page zoom map
  sigState.pageZoom = new Array(sigState.pages.length).fill(1);
  sigState.pages.forEach((_, i) => applyPageZoom(i));
}

// Render the given PDF bytes into the modal doc area (reused by rotate)
async function renderPdfPages(pdfBytes) {
  const modalDoc = document.getElementById('modalDoc');
  const pdf = await pdfjsLib.getDocument({ data: pdfBytes.slice() }).promise;
  modalDoc.innerHTML = ''; sigState.pages = []; sigState.pageInfo = []; sigState.pageWraps = [];
  // Base fit-to-width scale
  const containerWidth = Math.max(320, modalDoc.clientWidth - 40);
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const raw = page.getViewport({ scale: 1 });
    const scale = Math.min(1.5, containerWidth / raw.width);
    const viewport = page.getViewport({ scale });
    const sizer = document.createElement('div');
    sizer.className = 'pdf-page-wrap';
    sizer.style.width = viewport.width + 'px';
    sizer.style.height = viewport.height + 'px';
    // Per-page zoom controls
    const pageIdx = i - 1;
    const controls = document.createElement('div');
    controls.className = 'page-zoom-controls';
    controls.innerHTML = `
      <button class="page-num-btn" onclick="focusPage(${pageIdx})">📄 Page ${i}</button>
      <button onclick="pageZoomOut(${pageIdx})">➖</button>
      <span class="page-zoom-badge">100%</span>
      <button onclick="pageZoomIn(${pageIdx})">➕</button>
      <button onclick="rotatePage(${pageIdx})" title="Rotate 90°">🔄</button>
    `;
    sizer.appendChild(controls);
    const wrap = document.createElement('div');
    wrap.className = 'pdf-page';
    wrap.style.width = viewport.width + 'px'; wrap.style.height = viewport.height + 'px';
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width; canvas.height = viewport.height;
    wrap.appendChild(canvas);
    sizer.appendChild(wrap);
    modalDoc.appendChild(sizer);
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    sigState.pages.push(wrap);
    sigState.pageWraps.push(sizer);
    sigState.pageInfo.push({ w: viewport.width, h: viewport.height });
  }
  // Add a page navigator strip at the top
  if (sigState.pages.length > 1) {
    const nav = document.createElement('div');
    nav.className = 'page-navigator';
    nav.innerHTML = `
      <div class="page-nav-label">📚 ${sigState.pages.length} pages - click a page to focus on it:</div>
      <div class="page-nav-buttons">
        ${sigState.pages.map((_, i) => `<button onclick="focusPage(${i})">📄 Page ${i + 1}</button>`).join('')}
      </div>
    `;
    modalDoc.insertBefore(nav, modalDoc.firstChild);
  }
  sigState.pageWraps.forEach((sizer, i) => attachPinchZoomOnPage(sizer, i));
}

function applyPageZoom(idx) {
  const z = pageZoom(idx);
  const page = sigState.pages[idx];
  const sizer = sigState.pageWraps[idx];
  if (!page || !sizer) return;
  page.style.transform = `scale(${z})`;
  sizer.style.width = (page.offsetWidth * z) + 'px';
  sizer.style.height = (page.offsetHeight * z) + 'px';
  const badge = sizer.querySelector('.page-zoom-badge');
  if (badge) badge.textContent = Math.round(z * 100) + '%';
}

function attachPinchZoomOnPage(sizer, idx) {
  let pinch = null;
  const dist = (t1, t2) => Math.hypot(t1.clientX - t2.clientX, t1.clientY - t2.clientY);
  sizer.addEventListener('touchstart', (e) => {
    if (e.touches.length === 2) {
      pinch = { d0: dist(e.touches[0], e.touches[1]), z0: pageZoom(idx) };
      e.preventDefault();
    }
  }, { passive: false });
  sizer.addEventListener('touchmove', (e) => {
    if (e.touches.length === 2 && pinch) {
      const d = dist(e.touches[0], e.touches[1]);
      const nz = Math.max(0.5, Math.min(4, pinch.z0 * (d / pinch.d0)));
      sigState.pageZoom[idx] = nz;
      applyPageZoom(idx);
      e.preventDefault();
    }
  }, { passive: false });
  sizer.addEventListener('touchend', () => { pinch = null; });
  sizer.addEventListener('wheel', (e) => {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      const delta = -e.deltaY * 0.005;
      sigState.pageZoom[idx] = Math.max(0.5, Math.min(4, pageZoom(idx) + delta));
      applyPageZoom(idx);
    }
  }, { passive: false });
}

function zoomFit() {
  if (!sigState.pageZoom) return;
  sigState.pageZoom = sigState.pageZoom.map(() => 1);
  sigState.pages.forEach((_, i) => applyPageZoom(i));
}

function focusPage(idx) {
  const modalDoc = document.getElementById('modalDoc');
  if (!modalDoc) return;
  modalDoc.classList.add('focus-mode');
  sigState.pageWraps.forEach((sizer, i) => {
    sizer.classList.toggle('focused', i === idx);
  });
  // Add banner
  let banner = document.getElementById('focusBanner');
  if (!banner) {
    banner = document.createElement('div');
    banner.id = 'focusBanner';
    banner.className = 'focus-banner';
    modalDoc.insertBefore(banner, modalDoc.firstChild);
  }
  const total = sigState.pages.length;
  banner.innerHTML = `
    <span>📄 Viewing Page ${idx + 1} of ${total}</span>
    <div style="display:flex;gap:6px;">
      ${idx > 0 ? `<button onclick="focusPage(${idx - 1})">← Prev</button>` : ''}
      ${idx < total - 1 ? `<button onclick="focusPage(${idx + 1})">Next →</button>` : ''}
      <button onclick="unfocusPages()">✖ Show all pages</button>
    </div>`;
  sigState.focusedPage = idx;
  // Scroll into view
  modalDoc.scrollTo({ top: 0, behavior: 'smooth' });
}

function unfocusPages() {
  const modalDoc = document.getElementById('modalDoc');
  if (modalDoc) modalDoc.classList.remove('focus-mode');
  (sigState.pageWraps || []).forEach(sizer => sizer.classList.remove('focused'));
  const banner = document.getElementById('focusBanner'); if (banner) banner.remove();
  sigState.focusedPage = null;
}
function pageZoomIn(idx) { sigState.pageZoom[idx] = Math.min(4, pageZoom(idx) + 0.25); applyPageZoom(idx); }
function pageZoomOut(idx) { sigState.pageZoom[idx] = Math.max(0.5, pageZoom(idx) - 0.25); applyPageZoom(idx); }

// Rotate one page 90 deg clockwise by re-drawing its content rotated into a new page.
// This bakes the rotation into the page content (no /Rotate flag), so all existing
// screen-to-PDF coordinate math keeps working unchanged.
async function rotatePage(idx) {
  if (sigState.overlays.some(o => o.pageIdx === idx)) {
    if (!confirm('You have unsaved items on this page. Rotating will remove them. Continue?')) return;
    sigState.overlays.filter(o => o.pageIdx === idx).forEach(o => o.el.remove());
    sigState.overlays = sigState.overlays.filter(o => o.pageIdx !== idx);
    updateOverlayList();
  }
  const wasFocused = (typeof sigState.focusedPage === 'number') ? sigState.focusedPage : null;
  const savedZoom = (sigState.pageZoom || []).slice();
  toast('Rotating page...');
  try {
    const src = await PDFLib.PDFDocument.load(sigState.pdfBytes);
    const out = await PDFLib.PDFDocument.create();
    const count = src.getPageCount();
    for (let i = 0; i < count; i++) {
      const srcPage = src.getPage(i);
      const w = srcPage.getWidth(), hgt = srcPage.getHeight();
      const [embedded] = await out.embedPages([srcPage]);
      if (i === idx) {
        const newPage = out.addPage([hgt, w]);
        newPage.drawPage(embedded, { x: 0, y: w, width: w, height: hgt, rotate: PDFLib.degrees(-90) });
      } else {
        const newPage = out.addPage([w, hgt]);
        newPage.drawPage(embedded, { x: 0, y: 0, width: w, height: hgt });
      }
    }
    sigState.pdfBytes = await out.save();
    // Re-render from the rotated bytes
    unfocusPages();
    const modalDoc = document.getElementById('modalDoc');
    modalDoc.classList.remove('focus-mode');
    await renderPdfPages(sigState.pdfBytes);
    sigState.pageZoom = sigState.pages.map((_, i) => savedZoom[i] || 1);
    sigState.pages.forEach((_, i) => applyPageZoom(i));
    if (wasFocused !== null && sigState.pages[wasFocused]) focusPage(wasFocused);
    toast('✓ Page rotated');
  } catch (err) {
    toast('Rotate failed: ' + (err.message || err), 'error');
  }
}

// Get zoom per page (defaults to 1)
function pageZoom(idx) { return sigState.pageZoom && sigState.pageZoom[idx] || 1; }

// Figure out which page has the biggest visible overlap
function visiblePageAndPos(ovW, ovH) {
  const container = document.getElementById('modalDoc');
  const containerRect = container.getBoundingClientRect();
  let bestIdx = 0, bestOverlap = -1;
  sigState.pages.forEach((p, i) => {
    const r = p.getBoundingClientRect();
    const overlap = Math.max(0, Math.min(r.bottom, containerRect.bottom) - Math.max(r.top, containerRect.top));
    if (overlap > bestOverlap) { bestOverlap = overlap; bestIdx = i; }
  });
  const pageWrap = sigState.pages[bestIdx];
  const pageRect = pageWrap.getBoundingClientRect();
  const visCX = containerRect.left + containerRect.width / 2;
  const visCY = containerRect.top + containerRect.height / 2;
  const z = pageZoom(bestIdx);
  const left = Math.max(10, Math.min(pageWrap.offsetWidth - ovW - 10, (visCX - pageRect.left) / z - ovW / 2));
  const top = Math.max(10, Math.min(pageWrap.offsetHeight - ovH - 10, (visCY - pageRect.top) / z - ovH / 2));
  return { pageIdx: bestIdx, pageWrap, left, top };
}

async function bakeSignaturesIntoPdf() {
  const pdfDoc = await PDFLib.PDFDocument.load(sigState.pdfBytes, { updateMetadata: false });
  const font = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
  const pdfPages = pdfDoc.getPages();
  const imgCache = new Map();
  for (const ov of sigState.overlays) {
    const info = sigState.pageInfo[ov.pageIdx];
    const pdfPage = pdfPages[ov.pageIdx];
    const pdfW = pdfPage.getWidth(), pdfH = pdfPage.getHeight();
    const scaleX = pdfW / info.w, scaleY = pdfH / info.h;
    const x = ov.el.offsetLeft * scaleX;
    const w = ov.el.offsetWidth * scaleX;
    const hh = ov.el.offsetHeight * scaleY;
    const y = pdfH - (ov.el.offsetTop * scaleY) - hh;
    if (ov.type === 'image') {
      let img = imgCache.get(ov.dataUrl);
      if (!img) {
        const b64 = ov.dataUrl.split(',')[1];
        const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        img = ov.dataUrl.startsWith('data:image/jpeg') ? await pdfDoc.embedJpg(bytes) : await pdfDoc.embedPng(bytes);
        imgCache.set(ov.dataUrl, img);
      }
      pdfPage.drawImage(img, { x, y, width: w, height: hh });
    } else {
      const fontSize = Math.max(12, hh * 0.55);
      if (ov.isStamp) {
        pdfPage.drawRectangle({ x, y, width: w, height: hh, borderColor: PDFLib.rgb(0.86, 0.15, 0.15), borderWidth: 2 });
        const textW = font.widthOfTextAtSize(ov.text, fontSize);
        pdfPage.drawText(ov.text, { x: x + (w - textW) / 2, y: y + (hh - fontSize) / 2 + 4, size: fontSize, font, color: PDFLib.rgb(0.86, 0.15, 0.15) });
      } else {
        pdfPage.drawText(ov.text, { x: x + 4, y: y + (hh - fontSize) / 2 + 2, size: fontSize, font, color: PDFLib.rgb(0.02, 0.31, 0.18) });
      }
    }
  }
  // useObjectStreams:false serializes noticeably faster (slightly larger file)
  return await pdfDoc.save({ useObjectStreams: false });
}

function toast(message, type = 'success') {
  document.querySelectorAll('.pgds-toast').forEach(t => t.remove());
  const t = document.createElement('div');
  t.className = 'pgds-toast ' + type;
  t.textContent = message;
  document.body.appendChild(t);
  setTimeout(() => { t.classList.add('hide'); setTimeout(() => t.remove(), 400); }, 3000);
}

// Lock placed overlays in place visually (they are now baked into the PDF)
function lockPlacedOverlays() {
  sigState.overlays.forEach(ov => {
    ov.el.classList.add('locked');
    ov.el.onmousedown = null;
    ov.el.ontouchstart = null;
    const rh = ov.el.querySelector('.resize-handle'); if (rh) rh.remove();
    const xh = ov.el.querySelector('.remove-handle'); if (xh) xh.remove();
  });
  sigState.overlays = [];
  updateOverlayList();
}

// Tracks the in-flight background upload so Approve can wait for it
let pendingUpload = null;

function setSaveIndicator(state) {
  const b = document.getElementById('saveBtn');
  if (!b) return;
  const ic = b.querySelector('.ic');
  if (state === 'uploading') { ic.textContent = '☁️'; b.style.opacity = '0.7'; }
  else if (state === 'done') { ic.textContent = '💾'; b.style.opacity = '1'; }
  else if (state === 'error') { ic.textContent = '⚠️'; b.style.opacity = '1'; }
}

async function uploadSignedVersion(signedBytes) {
  // Capture these now - the modal may be closed before the upload finishes
  const docId = sigState.doc.id;
  const prevPath = sigState.doc.signed_file_path;
  const signedPath = `${currentUser.id}/${Date.now()}-signed.pdf`;
  const up = await sb.storage.from('pgds-signed').upload(signedPath, signedBytes, { contentType: 'application/pdf' });
  if (up.error) throw up.error;
  const upd = await sb.from('pgds_documents').update({
    signed_file_path: signedPath, signed_by: currentUser.id
  }).eq('id', docId);
  if (upd.error) throw upd.error;
  if (sigState.doc && sigState.doc.id === docId) sigState.doc.signed_file_path = signedPath;
  // Clean up the superseded file without blocking
  if (prevPath) { sb.storage.from('pgds-signed').remove([prevPath]).catch(() => {}); }
  return signedPath;
}

async function saveSignatures() {
  if (!sigState.overlays.length) return toast('Add a signature, text, or stamp first.', 'warn');
  const btn = document.getElementById('saveBtn'); btn.disabled = true;
  try {
    // Bake locally - this is the only part the user has to wait for
    const signedBytes = await bakeSignaturesIntoPdf();
    sigState.pdfBytes = signedBytes;
    lockPlacedOverlays();
    toast('✓ Saved');
    btn.disabled = false;
    // Upload in the background so the user can keep signing immediately
    setSaveIndicator('uploading');
    pendingUpload = uploadSignedVersion(signedBytes)
      .then(() => { setSaveIndicator('done'); pendingUpload = null; })
      .catch(err => { setSaveIndicator('error'); toast('Upload failed: ' + (err.message || err), 'error'); pendingUpload = null; throw err; });
  } catch (err) {
    toast('Error: ' + (err.message || err), 'error');
    btn.disabled = false;
    setSaveIndicator('error');
  }
}

async function approveDoc() {
  const hasUnsaved = sigState.overlays.length > 0;
  const hasSigned = !!sigState.doc.signed_file_path;
  const alreadyApproved = sigState.doc.status === 'approved';
  if (!hasSigned && !hasUnsaved) {
    return toast('Add a signature and Save before approving.', 'warn');
  }
  let confirmMsg = alreadyApproved
    ? 'Save changes to this already-approved document?'
    : 'Approve this document and move to the Approved folder?';
  if (hasUnsaved) confirmMsg = alreadyApproved
    ? 'You have new unsaved signatures. Save changes to this approved document?'
    : 'You have unsaved signatures on the doc. Save AND approve now?';
  if (!confirm(confirmMsg)) return;
  const btn = document.getElementById('approveBtn'); btn.disabled = true; btn.querySelector('.ic').textContent = '⏳';
  try {
    // Wait for any background save that is still uploading
    if (pendingUpload) { try { await pendingUpload; } catch (e) {} }
    // If there are pending overlays, bake and upload them now
    if (hasUnsaved) {
      const signedBytes = await bakeSignaturesIntoPdf();
      sigState.pdfBytes = signedBytes;
      await uploadSignedVersion(signedBytes);
    }
    const upd = await sb.from('pgds_documents').update({
      status: 'approved',
      signed_file_path: sigState.doc.signed_file_path,
      signed_by: currentUser.id,
      signed_at: new Date().toISOString()
    }).eq('id', sigState.doc.id);
    if (upd.error) throw upd.error;
    closeSigningModal();
    renderAdminView(); updateAdminCounts();
    toast('Document approved and moved to the Approved folder.');
  } catch (err) {
    toast('Error: ' + (err.message || err), 'error');
    const b = document.getElementById('approveBtn');
    if (b) { b.disabled = false; b.querySelector('.ic').textContent = '✅'; }
  }
}

function closeSigningModal() {
  const m = document.getElementById('sigModal'); if (m) m.remove();
  sigState = { doc: null, pages: [], overlays: [], mySignatures: [], pdfBytes: null, pageInfo: [] };
}

async function logout() { await sb.auth.signOut(); currentUser = null; route(); }
function refreshApp() { window.location.reload(); }

// Auto-wrap every table in a horizontal-scroll container so mobile users can pan across columns
function wrapAllTables() {
  document.querySelectorAll('table').forEach(t => {
    if (t.parentElement && t.parentElement.classList.contains('table-scroll')) return;
    const w = document.createElement('div');
    w.className = 'table-scroll';
    t.parentNode.insertBefore(w, t);
    w.appendChild(t);
  });
}
const tableObserver = new MutationObserver(() => wrapAllTables());
tableObserver.observe(document.body, { childList: true, subtree: true });

route();
</script>
</body>
</html>

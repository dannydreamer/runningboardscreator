'use client';
import { useState, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';

export default function Home() {
  const [groupName, setGroupName] = useState('');
  const [namesText, setNamesText] = useState('');
  const [files, setFiles] = useState([]);
  const [status, setStatus] = useState(null);
  const [history, setHistory] = useState(() => {
    if (typeof window === 'undefined') return {};
    try { return JSON.parse(localStorage.getItem('workshopHistory') || '{}'); } catch { return {}; }
  });
  const [dragging, setDragging] = useState(false);
  const fileRef = useRef();

  const saveHistory = (name, names) => {
    const next = { ...history, [name]: names };
    setHistory(next);
    localStorage.setItem('workshopHistory', JSON.stringify(next));
  };

  const mergeGroup = (name) => {
    setGroupName(name);
    const existing = history[name].join('\n');
    setNamesText(t => t.trim() ? existing + '\n' + t : existing);
    setStatus({ type: 'info', msg: `נטענו ${history[name].length} שמות קיימים. הוסיפי שמות חדשים ולחצי "צרי קובץ".` });
  };

  const deleteGroup = (name) => {
    if (!confirm(`האם את בטוחה שברצונך למחוק את הקבוצה "${name}"?`)) return;
    const next = { ...history };
    delete next[name];
    setHistory(next);
    localStorage.setItem('workshopHistory', JSON.stringify(next));
  };

  const addFiles = (newFiles) => {
    setFiles(prev => {
      const existing = new Set(prev.map(f => f.name));
      return [...prev, ...[...newFiles].filter(f => !existing.has(f.name))];
    });
  };

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false);
    addFiles(e.dataTransfer.files);
  }, []);

  const readFileAsBase64 = (file) => new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = e => res(e.target.result);
    fr.onerror = rej;
    fr.readAsDataURL(file);
  });

  const readExcel = (file) => new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        let text = '';
        wb.SheetNames.forEach(sn => {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1 });
          rows.forEach(row => { if (row.length) text += row.join(' ') + '\n'; });
        });
        res(text);
      } catch(err) { rej(err); }
    };
    fr.readAsArrayBuffer(file);
  });

  const generate = async () => {
    if (!groupName.trim()) {
      setStatus({ type: 'error', msg: 'לאיזו קבוצה השמות האלו שייכים?' });
      return;
    }

    setStatus({ type: 'loading', msg: 'מעבד קבצים...' });

    try {
      const images = [];
      const texts = [namesText];

      for (const file of files) {
        if (file.type.startsWith('image/')) {
          images.push(await readFileAsBase64(file));
        } else if (file.name.match(/\.(xlsx|xls|csv)$/i)) {
          texts.push(await readExcel(file));
        } else {
          texts.push(await file.text());
        }
      }

      if (!namesText.trim() && images.length === 0 && texts.filter(Boolean).length === 0) {
        setStatus({ type: 'error', msg: 'לא הוזנו שמות או קבצים.' });
        return;
      }

      setStatus({ type: 'loading', msg: 'מזהה שמות...' });

      const extractRes = await fetch('/api/extract-names', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ images, texts: texts.filter(Boolean) })
      });

      if (!extractRes.ok) throw new Error('שגיאה בחילוץ שמות');
      const { names: rawNames, error } = await extractRes.json();
      if (error) throw new Error(error);

      let allNames = rawNames || [];

      if (history[groupName.trim()]) {
        allNames = [...new Set([...history[groupName.trim()], ...allNames])];
      }

      allNames = [...new Set(allNames)].filter(Boolean).sort((a, b) => a.localeCompare(b, 'he'));

      if (allNames.length === 0) {
        setStatus({ type: 'error', msg: 'לא נמצאו שמות בחומר שהוזן.' });
        return;
      }

      saveHistory(groupName.trim(), allNames);
      setStatus({ type: 'loading', msg: `יוצר קובץ Word עם ${allNames.length} שמות...` });

      const docRes = await fetch('/api/generate-docx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ groupName: groupName.trim(), names: allNames })
      });

      if (!docRes.ok) throw new Error('שגיאה ביצירת קובץ');

      const blob = await docRes.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${groupName.trim()}.docx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 3000);

      setStatus({ type: 'success', msg: `✓ קובץ נוצר בהצלחה עם ${allNames.length} שמות!` });
    } catch (e) {
      setStatus({ type: 'error', msg: `שגיאה: ${e.message}` });
    }
  };

  const statusColors = {
    loading: { bg: '#EEF4FF', color: '#1a56db', border: '#93c5fd' },
    error:   { bg: '#FEF2F2', color: '#b91c1c', border: '#fca5a5' },
    success: { bg: '#F0FDF4', color: '#15803d', border: '#86efac' },
    info:    { bg: '#FFFBEB', color: '#92400e', border: '#fcd34d' },
  };

  const sc = status ? statusColors[status.type] : null;

  return (
    <div dir="rtl" style={{ fontFamily: 'Arial, sans-serif', minHeight: '100vh', background: '#F8F7F4', padding: '24px 16px' }}>
      <div style={{ maxWidth: 640, margin: '0 auto' }}>

        <div style={{ marginBottom: 28, textAlign: 'center' }}>
          <img src="/logo.png" alt="לוגו" style={{ height: 80, marginBottom: 12 }} />
          <h1 style={{ fontSize: 22, fontWeight: 600, color: '#1a1a1a', margin: 0 }}>📋 נוכחות סדנאות</h1>
          <p style={{ fontSize: 13, color: '#888', marginTop: 4 }}>הוסיפי שמות, קבלי קובץ Word מסודר</p>
        </div>

        {/* Group name */}
        <div style={cardStyle}>
          <label style={labelStyle}>שם הקבוצה / הסדנה</label>
          <input
            value={groupName}
            onChange={e => setGroupName(e.target.value)}
            placeholder="לדוגמה: קבוצת מנהלים — מרץ 2026"
            style={inputStyle}
          />
        </div>

        {/* Names input */}
        <div style={cardStyle}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
            <label style={{ ...labelStyle, marginBottom: 0 }}>שמות המשתתפים</label>
            {namesText && (
              <button onClick={() => setNamesText('')}
                style={{ background: 'none', border: '0.5px solid #d1d5db', borderRadius: 6, fontSize: 11, color: '#aaa', cursor: 'pointer', padding: '2px 8px' }}>
                נקה ✕
              </button>
            )}
          </div>
          <textarea
            value={namesText}
            onChange={e => setNamesText(e.target.value)}
            placeholder={'ישראל ישראלי\nשרה כהן\nמיכאל לוי'}
            style={{ ...inputStyle, minHeight: 110, resize: 'vertical', lineHeight: 1.7 }}
          />

          {/* File drop zone */}
          <div
            onDragOver={e => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            onClick={() => fileRef.current.click()}
            style={{
              marginTop: 12,
              border: `2px dashed ${dragging ? '#6366f1' : '#d1d5db'}`,
              borderRadius: 10,
              padding: '18px 16px',
              textAlign: 'center',
              cursor: 'pointer',
              background: dragging ? '#EEF2FF' : 'transparent',
              transition: 'all 0.15s',
              color: '#888',
              fontSize: 13
            }}
          >
            <div style={{ fontSize: 22, marginBottom: 4 }}>📎</div>
            <div>גרירה או לחיצה להעלאת קבצים</div>
            <div style={{ fontSize: 11, marginTop: 3, color: '#aaa' }}>Excel, CSV, תמונות, טקסט</div>
          </div>
          <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv,image/*,.txt" style={{ display: 'none' }}
            onChange={e => addFiles(e.target.files)} />

          {files.length > 0 && (
            <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 8 }}>
              {files.map((f, i) => (
                <div key={i} style={{ background: '#f3f4f6', border: '0.5px solid #e5e7eb', borderRadius: 8, padding: '4px 10px', fontSize: 12, color: '#555', display: 'flex', alignItems: 'center', gap: 6 }}>
                  {f.name}
                  <button onClick={() => setFiles(prev => prev.filter((_, j) => j !== i))}
                    style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#aaa', fontSize: 14, padding: 0, lineHeight: 1 }}>×</button>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* History */}
        {Object.keys(history).length > 0 && (
          <div style={cardStyle}>
            <div style={{ fontSize: 11, fontWeight: 600, color: '#aaa', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 10 }}>קבוצות קיימות — הוספת שמות</div>
            {Object.keys(history).map(name => (
              <div key={name} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 0', borderBottom: '0.5px solid #f0f0f0' }}>
                <div>
                  <div style={{ fontSize: 14, color: '#222' }}>{name}</div>
                  <div style={{ fontSize: 12, color: '#aaa' }}>{history[name].length} שמות</div>
                </div>
                <div style={{ display: 'flex', gap: 8 }}>
                  <button onClick={() => mergeGroup(name)}
                    style={{ padding: '5px 12px', background: 'none', border: '0.5px solid #d1d5db', borderRadius: 8, fontSize: 12, cursor: 'pointer', color: '#555' }}>
                    הוסף שמות ↗
                  </button>
                  <button onClick={() => deleteGroup(name)}
                    style={{ padding: '5px 10px', background: 'none', border: '0.5px solid #fca5a5', borderRadius: 8, fontSize: 12, cursor: 'pointer', color: '#ef4444' }}>
                    מחק
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}

        <button onClick={generate} style={{ width: '100%', padding: '14px', background: '#1a1a1a', color: '#fff', border: 'none', borderRadius: 10, fontSize: 15, fontWeight: 500, cursor: 'pointer', fontFamily: 'Arial, sans-serif' }}>
          📄 צרי קובץ Word
        </button>

        {status && (
          <div style={{ marginTop: 16, padding: '14px 16px', borderRadius: 10, fontSize: 14, background: sc.bg, color: sc.color, border: `1px solid ${sc.border}` }}>
            {status.type === 'loading' && <span style={{ marginLeft: 8 }}>⏳</span>}
            {status.msg}
          </div>
        )}

      </div>
    </div>
  );
}

const cardStyle = { background: '#fff', border: '0.5px solid #e5e7eb', borderRadius: 12, padding: 20, marginBottom: 16 };
const labelStyle = { fontSize: 13, fontWeight: 500, color: '#666', display: 'block', marginBottom: 6 };
const inputStyle = { width: '100%', padding: '10px 12px', fontSize: 15, border: '0.5px solid #d1d5db', borderRadius: 8, background: '#fff', color: '#1a1a1a', fontFamily: 'Arial, sans-serif', direction: 'rtl', outline: 'none' };

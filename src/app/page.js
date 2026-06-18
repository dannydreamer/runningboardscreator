'use client';
import { useState, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

const ROOM_CAPS = { 1: 20, 2: 12, 3: 16 };

const COLORS = ['לבן', 'אדום', 'כחול', 'שחור', 'ורוד', 'ירוק', 'סגול'];
const FACILITATORS = [
  'איילת השחר לוי',
  'אביטל לוי כץ',
  'דקלה אורן',
  'דניאל פראג',
  'שהירה הייקל',
  'מיקה מרים גובר',
  'מיה בודאי',
  'רועי הרץ',
];

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
  const [selectedRooms, setSelectedRooms] = useState(new Set());
  const [roomDetails, setRoomDetails] = useState({ 1: { color: '', facilitator: '' }, 2: { color: '', facilitator: '' }, 3: { color: '', facilitator: '' } });
  const [pendingRoomData, setPendingRoomData] = useState(null);
  const fileRef = useRef();

  const toggleRoom = (r) => {
    setSelectedRooms(prev => {
      const next = new Set(prev);
      if (next.has(r)) {
        next.delete(r);
        setRoomDetails(d => ({ ...d, [r]: { color: '', facilitator: '' } }));
      } else {
        next.add(r);
      }
      return next;
    });
  };

  const setRoomDetail = (r, field, value) => {
    setRoomDetails(d => ({ ...d, [r]: { ...d[r], [field]: value } }));
  };

  const usedFacilitators = (excludeRoom) =>
    Object.entries(roomDetails)
      .filter(([r, d]) => Number(r) !== excludeRoom && selectedRooms.has(Number(r)) && d.facilitator)
      .map(([, d]) => d.facilitator);

  const distributeToRooms = (names, rooms) => {
    const sorted = [...rooms].sort((a, b) => a - b);
    const assigned = {};
    sorted.forEach(r => assigned[r] = []);

    let remaining = [...names];
    let pool = [...sorted];

    while (remaining.length > 0 && pool.length > 0) {
      const k = pool.length;
      const n = remaining.length;
      const base = Math.floor(n / k);
      const extra = n % k;

      let idx = 0;
      const nextPool = [];
      const leftover = [];

      for (let i = 0; i < pool.length; i++) {
        const room = pool[i];
        const alloc = base + (i < extra ? 1 : 0);
        const cap = ROOM_CAPS[room];
        const space = cap - assigned[room].length;

        if (alloc <= space) {
          assigned[room].push(...remaining.slice(idx, idx + alloc));
          idx += alloc;
          if (assigned[room].length < cap) nextPool.push(room);
        } else {
          assigned[room].push(...remaining.slice(idx, idx + space));
          idx += space;
          leftover.push(...remaining.slice(idx, idx + (alloc - space)));
          idx += (alloc - space);
        }
      }

      remaining = leftover;
      pool = nextPool;
    }

    const overflow = {};
    sorted.forEach(r => overflow[r] = []);

    if (remaining.length > 0) {
      const overflowRooms = sorted.filter(r => r !== 2);
      for (let i = 0; i < remaining.length; i++) {
        overflow[overflowRooms[i % overflowRooms.length]].push(remaining[i]);
      }
    }

    return { assigned, overflow };
  };

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

  const readDocx = (file) => new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = async e => {
      try {
        const result = await mammoth.extractRawText({ arrayBuffer: e.target.result });
        res(result.value);
      } catch(err) { rej(err); }
    };
    fr.readAsArrayBuffer(file);
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

  const buildDocx = async (gName, allNames, roomData) => {
    setStatus({ type: 'loading', msg: `יוצר קובץ Word עם ${allNames.length} שמות...` });
    const body = roomData
      ? { groupName: gName, names: allNames, rooms: roomData }
      : { groupName: gName, names: allNames };

    const docRes = await fetch('/api/generate-docx', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });

    if (!docRes.ok) throw new Error('שגיאה ביצירת קובץ');
    const blob = await docRes.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${gName}.docx`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    setStatus({ type: 'success', msg: `✓ קובץ נוצר בהצלחה עם ${allNames.length} שמות!` });
  };

  const generate = async () => {
    if (!groupName.trim()) {
      setStatus({ type: 'error', msg: 'לאיזו קבוצה השמות האלו שייכים?' });
      return;
    }

    setStatus({ type: 'loading', msg: 'מעבד קבצים...' });
    setPendingRoomData(null);

    try {
      const images = [];
      const texts = [namesText];

      for (const file of files) {
        if (file.type.startsWith('image/')) {
          images.push(await readFileAsBase64(file));
        } else if (file.name.match(/\.(xlsx|xls|csv)$/i)) {
          texts.push(await readExcel(file));
        } else if (file.name.match(/\.docx$/i)) {
          texts.push(await readDocx(file));
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

      // No rooms selected → current behavior
      if (selectedRooms.size === 0) {
        await buildDocx(groupName.trim(), allNames, null);
        return;
      }

      // Validate: only Room 2 selected and count > 12
      if (selectedRooms.size === 1 && selectedRooms.has(2) && allNames.length > 12) {
        setStatus({ type: 'error', msg: 'חדר 2 מוגבל ל-12 משתתפים. יש לבחור גם חדר 1 ו/או חדר 3 כדי להמשיך.' });
        return;
      }

      const { assigned, overflow } = distributeToRooms(allNames, [...selectedRooms]);

      const roomData = [...selectedRooms].sort((a, b) => a - b).map(r => ({
        number: r,
        names: assigned[r],
        overflowNames: overflow[r],
        color: roomDetails[r].color,
        facilitator: roomDetails[r].facilitator,
      }));

      const overflowWarnings = roomData
        .filter(rd => rd.overflowNames.length > 0)
        .map(rd => `חדר ${rd.number} חצה את הקיבולת המקסימלית ב-${rd.overflowNames.length} משתתפים`);

      if (overflowWarnings.length > 0) {
        setStatus(null);
        setPendingRoomData({ groupName: groupName.trim(), allNames, rooms: roomData, overflowWarnings });
        return;
      }

      await buildDocx(groupName.trim(), allNames, roomData);
    } catch (e) {
      setStatus({ type: 'error', msg: `שגיאה: ${e.message}` });
    }
  };

  const confirmAndGenerate = async () => {
    const { groupName: gName, allNames, rooms } = pendingRoomData;
    setPendingRoomData(null);
    try {
      await buildDocx(gName, allNames, rooms);
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
          <h1 style={{ fontSize: 22, fontWeight: 600, color: '#1a1a1a', margin: 0 }}>לוחות רצים</h1>
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
            <div style={{ fontSize: 11, marginTop: 3, color: '#aaa' }}>Excel, CSV, Word, תמונות, טקסט</div>
          </div>
          <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv,.docx,image/*,.txt" style={{ display: 'none' }}
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

        {/* Room selection */}
        <div style={cardStyle}>
          <label style={labelStyle}>חדרים פעילים</label>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
            {[1, 2, 3].map(r => {
              const active = selectedRooms.has(r);
              return (
                <div key={r} onClick={() => toggleRoom(r)} style={{
                  display: 'inline-flex', alignItems: 'center', gap: 8,
                  padding: '9px 18px', borderRadius: 24,
                  background: active ? '#1a1a1a' : '#fff',
                  color: active ? '#fff' : '#444',
                  border: `1.5px solid ${active ? '#1a1a1a' : '#d1d5db'}`,
                  cursor: 'pointer', fontSize: 14, fontFamily: 'Arial, sans-serif',
                  userSelect: 'none', transition: 'all 0.18s',
                  boxShadow: active ? '0 2px 8px rgba(0,0,0,0.15)' : 'none'
                }}>
                  <span style={{
                    width: 16, height: 16, borderRadius: 4,
                    border: `2px solid ${active ? '#fff' : '#c0c0c0'}`,
                    background: active ? '#fff' : 'transparent',
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    fontSize: 10, color: '#1a1a1a', flexShrink: 0, transition: 'all 0.18s'
                  }}>{active ? '✓' : ''}</span>
                  חדר {r}
                  <span style={{ fontSize: 11, opacity: 0.55 }}>עד {ROOM_CAPS[r]}</span>
                </div>
              );
            })}
          </div>
          {selectedRooms.size === 0 && (
            <p style={{ fontSize: 12, color: '#bbb', marginTop: 10, marginBottom: 0 }}>ללא בחירת חדרים — הקובץ ייוצר ללא חלוקה לחדרים</p>
          )}

          {[1, 2, 3].filter(r => selectedRooms.has(r)).map(r => {
            const taken = usedFacilitators(r);
            return (
              <div key={r} style={{ marginTop: 14, padding: '12px 14px', background: '#F9FAFB', borderRadius: 10, border: '0.5px solid #e5e7eb' }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: '#555', marginBottom: 10 }}>חדר {r}</div>
                <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
                  <div style={{ flex: 1, minWidth: 120 }}>
                    <label style={{ fontSize: 11, color: '#aaa', display: 'block', marginBottom: 4 }}>צבע</label>
                    <select
                      value={roomDetails[r].color}
                      onChange={e => setRoomDetail(r, 'color', e.target.value)}
                      style={selectStyle}
                    >
                      <option value="">— בחרי צבע —</option>
                      {COLORS.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  <div style={{ flex: 2, minWidth: 180 }}>
                    <label style={{ fontSize: 11, color: '#aaa', display: 'block', marginBottom: 4 }}>מנחה</label>
                    <select
                      value={roomDetails[r].facilitator}
                      onChange={e => setRoomDetail(r, 'facilitator', e.target.value)}
                      style={selectStyle}
                    >
                      <option value="">— בחרי מנחה —</option>
                      {FACILITATORS.filter(f => !taken.includes(f)).map(f => (
                        <option key={f} value={f}>{f}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        <button onClick={generate} style={{ width: '100%', padding: '14px', background: '#1a1a1a', color: '#fff', border: 'none', borderRadius: 10, fontSize: 15, fontWeight: 500, cursor: 'pointer', fontFamily: 'Arial, sans-serif' }}>
          📄 צרי קובץ Word
        </button>

        {status && (
          <div style={{ marginTop: 16, padding: '14px 16px', borderRadius: 10, fontSize: 14, background: sc.bg, color: sc.color, border: `1px solid ${sc.border}` }}>
            {status.type === 'loading' && <span style={{ marginLeft: 8 }}>⏳</span>}
            {status.msg}
          </div>
        )}

        {pendingRoomData && (
          <div style={{ marginTop: 16, padding: '16px 18px', borderRadius: 10, background: '#FFFBEB', border: '1px solid #fcd34d', color: '#92400e' }}>
            {pendingRoomData.overflowWarnings.map((w, i) => (
              <div key={i} style={{ fontSize: 14, marginBottom: 4 }}>
                ⚠️ אזהרה: {w}. השמות החורגים יוצגו בצבע אדום במסמך.
              </div>
            ))}
            <div style={{ fontSize: 13, color: '#78350f', marginTop: 6 }}>אפשר להמשיך וליצור את המסמך כרגיל.</div>
            <div style={{ display: 'flex', gap: 10, marginTop: 14 }}>
              <button onClick={confirmAndGenerate} style={{ padding: '9px 20px', background: '#1a1a1a', color: '#fff', border: 'none', borderRadius: 8, fontSize: 14, cursor: 'pointer', fontFamily: 'Arial, sans-serif' }}>
                המשיכי וצרי קובץ ←
              </button>
              <button onClick={() => setPendingRoomData(null)} style={{ padding: '9px 16px', background: 'none', border: '1px solid #e5c54b', borderRadius: 8, fontSize: 14, cursor: 'pointer', color: '#92400e', fontFamily: 'Arial, sans-serif' }}>
                ביטול
              </button>
            </div>
          </div>
        )}

        <div style={{ marginTop: 40, textAlign: 'center' }}>
          <img src="/funny.png" alt="" style={{ width: '100%', maxWidth: 600, borderRadius: 12 }} />
        </div>

      </div>
    </div>
  );
}

const cardStyle = { background: '#fff', border: '0.5px solid #e5e7eb', borderRadius: 12, padding: 20, marginBottom: 16 };
const labelStyle = { fontSize: 13, fontWeight: 500, color: '#666', display: 'block', marginBottom: 6 };
const inputStyle = { width: '100%', padding: '10px 12px', fontSize: 15, border: '0.5px solid #d1d5db', borderRadius: 8, background: '#fff', color: '#1a1a1a', fontFamily: 'Arial, sans-serif', direction: 'rtl', outline: 'none' };
const selectStyle = { width: '100%', padding: '8px 10px', fontSize: 13, border: '0.5px solid #d1d5db', borderRadius: 8, background: '#fff', color: '#1a1a1a', fontFamily: 'Arial, sans-serif', direction: 'rtl', outline: 'none', cursor: 'pointer' };

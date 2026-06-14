import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle
} from 'docx';

export async function POST(req) {
  try {
    const { groupName, names, rooms } = await req.json();

    const pageW = 11906, margin = 1000;
    const contentW = pageW - margin * 2;

    const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
    const borders = { top: border, bottom: border, left: border, right: border };

    let table;

    if (rooms && rooms.length > 0) {
      // Room-based layout. Room 1 on the right (RTL) = last column in LTR array.
      // So sort descending: [3, 2, 1] left→right, which puts Room 1 on the right.
      const orderedRooms = [...rooms].sort((a, b) => b.number - a.number);
      const numCols = orderedRooms.length;
      const colW = Math.floor(contentW / numCols);

      // Header row
      const headerCells = orderedRooms.map(room => new TableCell({
        width: { size: colW, type: WidthType.DXA },
        borders,
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        shading: { fill: 'F3F4F6', type: 'clear' },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `חדר ${room.number}`, font: 'Arial', size: 24, bold: true, color: '374151' })]
        })]
      }));

      const maxRows = Math.max(...orderedRooms.map(r => r.names.length + r.overflowNames.length));
      const contentRows = [];

      for (let row = 0; row < maxRows; row++) {
        const cells = orderedRooms.map(room => {
          const allRoomNames = [...room.names, ...room.overflowNames];
          const name = row < allRoomNames.length ? allRoomNames[row] : '';
          const isOverflow = row >= room.names.length && name !== '';

          return new TableCell({
            width: { size: colW, type: WidthType.DXA },
            borders,
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({
                text: name,
                font: 'Arial',
                size: 22,
                color: isOverflow ? 'DC2626' : '000000'
              })]
            })]
          });
        });
        contentRows.push(new TableRow({ children: cells }));
      }

      table = new Table({
        width: { size: contentW, type: WidthType.DXA },
        columnWidths: Array(numCols).fill(colW),
        rows: [new TableRow({ children: headerCells }), ...contentRows]
      });

    } else {
      // Original auto-column logic
      const numCols = names.length <= 30 ? 1 : names.length <= 60 ? 2 : names.length <= 90 ? 3 : 4;
      const colW = Math.floor(contentW / numCols);
      const rowsPerCol = Math.ceil(names.length / numCols);

      const tableRows = [];
      for (let r = 0; r < rowsPerCol; r++) {
        const cells = [];
        for (let c = 0; c < numCols; c++) {
          const idx = c * rowsPerCol + r;
          const name = idx < names.length ? names[idx] : '';
          cells.push(new TableCell({
            width: { size: colW, type: WidthType.DXA },
            borders,
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: name, font: 'Arial', size: 22 })]
            })]
          }));
        }
        tableRows.push(new TableRow({ children: cells }));
      }

      table = new Table({
        width: { size: contentW, type: WidthType.DXA },
        columnWidths: Array(numCols).fill(colW),
        rows: tableRows
      });
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: { width: pageW, height: 16838 },
            margin: { top: margin, bottom: margin, left: margin, right: margin }
          }
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
            children: [new TextRun({ text: groupName, bold: true, size: 36, font: 'Arial' })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
            children: [new TextRun({ text: `סה"כ משתתפים: ${names.length}`, size: 22, font: 'Arial', color: '666666' })]
          }),
          table
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    return new Response(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(groupName)}.docx"`
      }
    });
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}

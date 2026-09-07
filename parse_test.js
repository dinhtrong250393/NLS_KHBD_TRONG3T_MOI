const content = `
[BLUE]Hoạt động 5
| Bước | Nội dung |
|---|---|
| Chuyển giao | Giao bài |
| Thực hiện | Làm bài |
[/BLUE]
`;
let inTable = false;
let currentColor = "000000";
let xmlParagraphs = '';
const lines = content.split('\n').filter(l => l.trim() !== '');

for (let line of lines) {
  if (line.includes('[RED]')) currentColor = "C00000";
  if (line.includes('[BLUE]')) currentColor = "2F5496";
  
  let cleanLine = line.replace(/\[RED\]/g, '').replace(/\[\/RED\]/g, '').replace(/\[BLUE\]/g, '').replace(/\[\/BLUE\]/g, '').trim();
  cleanLine = cleanLine.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  
  if (cleanLine.startsWith('|') && cleanLine.endsWith('|')) {
      if (cleanLine.includes('---')) continue; 
      
      if (!inTable) {
          inTable = true;
          xmlParagraphs += `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblW w:w="5000" w:type="pct"/></w:tblPr>`;
      }
      
      xmlParagraphs += `<w:tr>`;
      const cells = cleanLine.substring(1, cleanLine.length - 1).split('|');
      for (let cell of cells) {
          let text = cell.trim();
          xmlParagraphs += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:rPr><w:color w:val="${currentColor}"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;
      }
      xmlParagraphs += `</w:tr>`;
  } else {
      if (inTable) {
          inTable = false;
          xmlParagraphs += `</w:tbl>`;
      }
      xmlParagraphs += `<w:p><w:r><w:rPr><w:color w:val="${currentColor}"/></w:rPr><w:t>${cleanLine}</w:t></w:r></w:p>`;
  }
  
  if (line.includes('[/RED]') || line.includes('[/BLUE]')) currentColor = "000000";
}
if (inTable) xmlParagraphs += `</w:tbl>`;
console.log(xmlParagraphs);
